[num,txt,raw]=xlsread('unadjust_price.xlsx');%���վ㦬�L��
[numc,txtc,rawc]=xlsread('div_cash.xlsx');%�{���ѧQ���
[nums,txts,raws]=xlsread('div_stock.xlsx');%�Ѳ��ѧQ���
num_adjust = num;% �̫��ˬd�վ�
N1 = length(num(1,:)); % ���X�ѻ���Ʀ��
n2 = length(num(:,1))+2;%���X�ѻ���ƦC��
N3 = min(numc(:,1))+1;%���X���v����ƦC��
for n = 1:N1 %�C�����q�Ѳ��v�@�p��
for j=2:2:8 % Ū�����v�������m
for i=3:n2
     if isequal(raw(i,1),rawc(j,n+1))>0 %���o��{���ѧQ���
         temp = j-1; %�{���ѧQ��m
         for k = i:n2
             num_adjust(k-2,n) = num_adjust(k-2,n) + (1-0.1425/100)*numc(temp,n);%�[�^�{���ѧQ-�������
         end
         clear temp
     end
     if isequal(raw(i,1),raws(j,n+1))>0 %���o��Ѳ��ѧQ���
         temp = j-1;%�Ѳ��ѧQ��m
         for k = i:n2
             num_adjust(k-2,n) = num_adjust(k-2,n)*(1+nums(temp,n)/10); %�Ѳ��ѧQ�٭�
         end
         clear temp N1 N2 N3
    end
end
end
end
%csvwrite('adjust.csv',num_adjust);%�s��csv��


 