[num,txt,raw]=xlsread('unadjust_price.xlsx');
[numc,txtc,rawc]=xlsread('div_cash.xlsx');
[nums,txts,raws]=xlsread('div_stock.xlsx');
num_adjust = num;
jc = 2;
js = 2;
n=1;
 for i=3:990
     if isequal(raw(i,n),rawc(jc,n+1))>0 %加回股利
         jc = jc+1; % 股利位置
         for k = i:990
             num_adjust(k-3,n) = num_adjust(k-3,n) + numc(jc,n);%加回股利
         end
         jc = jc+1;
         if jc >= 7
             break
         end
     end
 end
 num_adjust(:,1)-num(:,1)
 