[num,txt,raw]=xlsread('unadjust_price.xlsx');
[numc,txtc,rawc]=xlsread('div_cash.xlsx');
[nums,txts,raws]=xlsread('div_stock.xlsx');
num_adjust = num;
jc = 2;
js = 2;
n=1;
%for n = 1:1590 
for i=3:990
     if jc<=7
     if isequal(raw(i,n),rawc(jc,n+1))>0 %�[�^�ѧQ
         jc = jc+1; % �ѧQ��m
         for k = i:990
             num_adjust(k-3,n) = num_adjust(k-3,n) + numc(jc,n);%�[�^�ѧQ
         end
         jc = jc+1;
     end
     end
     if isequal(raw(i,n),raws(js,n+1))>0 %�Ѳ��ѧQ�٭�
         js = js+1;%�Ѳ��ѧQ��m
         for k = i:990
             num_adjust(k-3,n) = numadhust(k-3,n)*(1+nums(js,n)/10);
         end
         js = js+1;
         if js>=7
             break
         end
     end
end
%end
 num_adjust(:,1)-num(:,1)
 