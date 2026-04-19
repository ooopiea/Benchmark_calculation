%% 1. 加载数据
filename = 'balanced_ICIO.xlsx';
fprintf('正在从 %s 读取数据，请稍候...\n', filename);

% 假设第一行是列名，第一列是行名，数据从 B2 单元格开始
% 读取数值矩阵 (不包含表头)
% 如果你的 Excel 版本较旧或文件过大，建议先转为 .csv 格式
raw_data = readmatrix(filename, 'Range', 'B2'); 

%% 2. 定义维度参数 (根据图片)
G = 81;   % 国家数
N = 50;   % 每个国家的部门数
K = 6;    % 每个国家的最终需求类别 (486 / 81 = 6)
GN = G * N; % 4050
GK = G * K; % 486

%% 3. 矩阵切片 (根据图片结构提取各部分)
% 中间投入矩阵 Z: [GN x GN] (4050x4050)
Z_large = raw_data(1:GN, 1:GN);

% 最终需求矩阵 FD: [GN x GK] (4050x486)
FD_large = raw_data(1:GN, GN+1 : GN+GK);

% 总产出向量 X (右侧的 OUT 列): [GN x 1]
X_large = raw_data(1:GN, GN+GK+1);

% 净税收 TLS (下方的紫色块): 跨越 Z 和 FD 区域
TLS_Z_large = raw_data(GN+1, 1:GN);       % 位于 Z 下方的 TLS
TLS_FD_large = raw_data(GN+1, GN+1:GN+GK); % 位于 FD 下方的 TLS

% 增加值 VA (下方的粉色块): [1 x GN]
VA_large = raw_data(GN+2, 1:GN);

% 注意：raw_data(GN+3, :) 通常是底部加总校验用的 OUT 行

%% 4. 构建压缩算子 (Summation Operator)
% S 是一个 [N x GN] 的矩阵，用于将各国的相同部门加总
S = repmat(eye(N), 1, G); 

%% 5. 执行世界投入产出表 (WIOT) 压缩计算
fprintf('正在进行矩阵压缩计算...\n');

% 5.1 压缩中间投入 (World Z) -> [N x N]
Z_world = S * Z_large * S';

% 5.2 压缩最终需求 (World FD) -> [N x K]
% 先合并行(供应部门)
FD_temp = S * FD_large; 
% 再合并列(跨国合并相同类型的需求)
FD_world = zeros(N, K);
for k = 1:K
    idx = k : K : GK; % 提取所有国家第 k 类需求的列索引
    FD_world(:, k) = sum(FD_temp(:, idx), 2);
end

% 5.3 压缩增加值 VA 和 总产出 X
VA_world = VA_large * S'; % [1 x N]
X_world = S * X_large;    % [N x 1]

% 5.4 压缩税收 TLS
TLS_Z_world = TLS_Z_large * S'; % [1 x N]
TLS_FD_world = zeros(1, K);
for k = 1:K
    idx = k : K : GK;
    TLS_FD_world(1, k) = sum(TLS_FD_large(1, idx));
end

%% 6. 平衡检查
X_check = sum(Z_world, 2) + sum(FD_world, 2);
error = max(abs(X_world - X_check));
fprintf('压缩完成！行收支平衡误差为: %e\n', error);

%% 导出结果 
%% 1. 初始化完整矩阵
% 计算维度：行 = N (行业) + 1 (TLS) + 1 (VA) + 1 (OUT)
%          列 = N (行业) + K (最终需求类别) + 1 (OUT)
Total_Rows = N + 3;
Total_Cols = N + K + 1;
WIOT = zeros(Total_Rows, Total_Cols);

%% 2. 填充数据块

% 填充左上角：中间投入矩阵 (N x N)
WIOT(1:N, 1:N) = Z_world;

% 填充中上部：最终需求矩阵 (N x K)
WIOT(1:N, N+1 : N+K) = FD_world;

% 填充右上角：总产出列向量 (N x 1)
WIOT(1:N, N+K+1) = X_world;

% 填充税收行 (TLS)：包括中间投入部分和最终需求部分
WIOT(N+1, 1:N) = TLS_Z_world;
WIOT(N+1, N+1 : N+K) = TLS_FD_world;

% 填充增加值行 (VA)：(1 x N)
WIOT(N+2, 1:N) = VA_world;

% 填充底部总产出行 (OUT)：(1 x N)，通常等于 X_world 的转置
WIOT(N+3, 1:N) = X_world';

%% 3. 添加表头（可选，方便导出到 Excel）
% 创建行业名称列表 (如 Ind1, Ind2...)
Sector_Names = arrayfun(@(x) ['Industry_', num2str(x)], 1:N, 'UniformOutput', false);
FD_Names = {'HFCE', 'NPISH', 'GGFC', 'GFCF', 'INVNT', 'DPABR'};
Col_Headers = [Sector_Names, FD_Names, {'Total_Output'}];
Row_Headers = [Sector_Names, {'TLS', 'VA', 'Total_Output'}]';

% 将数值转为 Cell 数组以便加入表头
WIOT_Cell = num2cell(WIOT);
% 在上方插入列名，在左侧插入行名
Final_Table = [ [{'Labels'}, Col_Headers]; [Row_Headers, WIOT_Cell] ];

%% 4. 导出到 Excel
output_file = 'WIOT2022_USDPPP.xlsx';
writecell(Final_Table, output_file, 'Sheet', 'WIOT');

fprintf('完整的世界投入产出表已构建并保存至: %s\n', output_file);