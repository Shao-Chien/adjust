[num,txt,raw]=xlsread('unadjust_price.xlsx');
[numc,txtc,rawc]=xlsread('div_cash.xlsx');
[nums,txts,raws]=xlsread('div_stock.xlsx');
num_adjust = num;
%n=56;
for n = 1:1590 
for j=2:2:8
for i=3:990
     if isequal(raw(i,1),rawc(j,n+1))>0 %�[�^�ѧQ
         temp = j-1; % �ѧQ��m
         for k = i:990
             num_adjust(k-3,n) = num_adjust(k-3,n) + numc(temp,n);%�[�^�ѧQ
         end
         clear temp
     end
     if isequal(raw(i,1),raws(j,n+1))>0 %�Ѳ��ѧQ�٭�
         temp = j-1;%�Ѳ��ѧQ��m
         for k = i:990
             num_adjust(k-3,n) = num_adjust(k-3,n)*(1+nums(temp,n)/10);
         end
         clear temp
    end
end
end
end
csvwrite('adjust.csv',num_adjust);


 