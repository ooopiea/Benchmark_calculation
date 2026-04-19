%本代码基于映射矩阵将投入产出表映射为不同维度
clc;
clear;
%% 数据读取
% 路径定义（相对路径）
path_IOT = 'WIOT2022_USDPPP.xlsx';
path_AlignTable = 'Align_Table.xlsx';
path_Emission = 'Handmade_IEA_EDGAR_2022_world_CO2emission.xlsx';

IO_old = readmatrix(path_IOT, 'Sheet', 'WIOT', 'Range', 'B2:BE54');
Emission_old = readmatrix(path_Emission, 'Sheet', 'OECD.IEA,WORLDBIGCO2,1.0,filter', 'Range', 'X3:X36');    
T_matrix = readmatrix(path_AlignTable, 'Sheet', '50TO22_OECD-IEA', 'Range', 'D3:Y52');
T_matrix = fillmissing(T_matrix, 'constant', 0);
T_matrix_emission = readmatrix(path_AlignTable, 'Sheet', '34TO22_IEA-OECD', 'Range', 'E3:Z36');
T_matrix_emission = fillmissing(T_matrix_emission, 'constant', 0);

%% 数据转换
% 1. 参数定义
n_old = 50; % 原有行业数
n_new = 22; % 目标行业数
fd_cols = 6; % 最终需求列数 
va_rows = 3; % 增加值行数

% 2. 拆分原始表 IO_old
% Z_old: 中间投入矩阵 (50x50)
Z_old = IO_old(1:n_old, 1:n_old);

% F_old: 最终需求矩阵 (50x6)
F_old = IO_old(1:n_old, n_old+1 : n_old+fd_cols);

% VA_old: 增加值矩阵 (3x50)
VA_old = IO_old(n_old+1 : n_old+va_rows, 1:n_old);

% Others: 表格右下角的角块（通常为0或VA对FD的投入，5x6）
Corner_old = IO_old(n_old+1:end, n_old+1:end);

% 3. 构建转换矩阵 M (28x50)
% 注意：readmatrix读入的T_matrix如果是50x28，需要转置为28x50进行左乘
M = T_matrix'; 

% 4. 执行映射转换 (核心公式)
% 中间投入转换：M * Z * M'
Z_new = M * Z_old * M';

% 最终需求转换：M * F
F_new = M * F_old;

% 增加值转换：VA * M'
VA_new = VA_old * M';
%% 
% 5. 拼接新表 IO_new (33x37 维)
% 新表结构：[Z_new(28x28), F_new(28x9); VA_new(5x28), Corner_old(5x9)]
IO_new = [Z_new, F_new; VA_new, Corner_old];

%% 结果验证
fprintf('原始表维度: %d x %d\n', size(IO_old,1), size(IO_old,2));
fprintf('新表维度: %d x %d\n', size(IO_new,1), size(IO_new,2));

% 校验行和：新表的总产出应等于旧表总产出的映射
X_old = sum(Z_old, 2) + sum(F_old, 2);
X_new_calc = M * X_old;
X_new_real = sum(Z_new, 2) + sum(F_new, 2);

if max(abs(X_new_calc - X_new_real)) < 1
    disp('转换成功：行向总量平衡。');
else
    disp('警告：总量不平衡，请检查映射矩阵列和是否为1。');
end       

%% 转换碳排放向量
% 转换
Emission_new = T_matrix_emission' * Emission_old;
               
%  验证总量守恒
fprintf('原始 50 部门总排放: %.2f Mt\n', sum(Emission_old) );
fprintf('转换后 22 部门总排放: %.2f Mt\n', sum(Emission_new));

%% 输出结果到Excel
% 输出投入产出表
writematrix(IO_new, 'IO_emission_new.xlsx', 'Sheet', 'IO_new');

% 输出碳排放向量
writematrix(Emission_new, 'IO_emission_new.xlsx', 'Sheet', 'Emission_new');

disp('结果已成功输出到 IO_emission_new.xlsx');