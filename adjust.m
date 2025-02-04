[num,txt,raw]=xlsread('unadjust_price.xlsx');%未調整收盤價
[numc,txtc,rawc]=xlsread('div_cash.xlsx');%現金股利資料
[nums,txts,raws]=xlsread('div_stock.xlsx');%股票股利資料
num_adjust = num;% 最後檢查調整
col = length(num(1,:)); % 取出股價資料行數N
row = length(num(:,1))+2;%取出股價資料列數
N = min(numc(:,1))+1;%取出除權息資料列數
for n = 1:col %每間公司股票逐一計算
for j=2:2:N % 讀取除權息日期位置
for i=3:row
     if isequal(raw(i,1),rawc(j,n+1))>0 %比對發放現金股利日期
         temp = j-1; %現金股利位置
         for k = i:row
             num_adjust(k-2,n) = num_adjust(k-2,n) + numc(temp,n)*(1-0.1425/100);%加回現金股利-交易成本
         end
         clear temp
     end
     if isequal(raw(i,1),raws(j,n+1))>0 %比對發放股票股利日期
         temp = j-1;%股票股利位置
         for k = i:row
             num_adjust(k-2,n) = num_adjust(k-2,n)*(1+nums(temp,n)/10); %股票股利還原
         end
         clear temp col row N
    end
end
end
end
%csvwrite('adjust.csv',num_adjust);%存成csv檔


 