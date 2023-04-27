% num2letter：将列号转化为Excel表格中对应字母编号的函数
% 关键词：二十六进制
% 输入值 col_num: 目标列号，double型
% 输出值 col_letter：Excel表格中对应表示，Char型

% 示例1：第24列在Excel表格中表示为'X'列，令 colName1 = ExcelCol(24)
% 可得 colName1 = 'X'

% 示例2：查看第252231列在Excel表格对应列，令 colName2 = ExcelCol(252231)
% 可得 colName2 = 'NICE'

function col_letter = num2letter(col_num)

string = {'A','B','C','D','E','F',...
          'G','H','I','J','K','L',...
          'M','N','O','P','Q','R',...
          'S','T','U','V','W','X','Y','Z'};

col_letter = [];

while col_num>0

    m = mod(col_num,26);

    if m == 0
        m = 26;
    end

    col_letter = [string{m} col_letter];

    col_num = (col_num-m)/26;
end

end