clc;
clear;

%% 计算直接消耗系数矩阵
% 计算新的矩阵A，其中A(i,j) = X(i,j) / x(j);
Path = 'IO_emission_new.xlsx';
aggregated_table_Z_FD = readmatrix(Path, 'Sheet', 'IO_new', 'Range', 'A1:AB16');%这个部分是Z和FD的总和
aggregated_table_Industrial = readmatrix(Path, 'Sheet', 'IO_new', 'Range', 'A1:P16');%这个部分是工业的总和
% 前提是，这16行都是实业；不包含服务业等，以建筑业截止
Aggregated_table_sum = sum(aggregated_table_Z_FD, 2);
Aggregated_table_sum_inter = repmat(Aggregated_table_sum',16,1);
Direct_consumption_M = aggregated_table_Industrial./Aggregated_table_sum_inter;

%% 求列昂惕夫逆矩阵
A=Direct_consumption_M;
I=eye(size(A));
% 计算矩阵的逆
A1=I-A;
L = inv(A1);

%% 计算生产部门的碳排放强度
Emission = readmatrix(Path, 'Sheet', 'Emission_new', 'Range', 'A1:A16');
Emission_intensity = Emission./Aggregated_table_sum * 10^(6);%单位为tCO2

%% 计算各行业基准值
baseline_M = Emission_intensity' * L;%tCO2/百万美元
%% 输出结果到Excel
% 先检查baseline_M的内容
disp('基准值矩阵大小:');
disp(size(baseline_M));
disp('基准值矩阵内容:');
disp(baseline_M);

writematrix(baseline_M, 'benchmark2022.xlsx', 'Sheet', 'benchmark2022', 'Range', 'A1');

disp('结果已成功输出到 benchmark2022.xlsx');


