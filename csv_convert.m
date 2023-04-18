%---------------------------------------------------------------------------
% приведение записей с системы индицирования ХНПЛ к "удобному" формату
% а: импортировать данные: Home -> import data -> numeric matrix (с именем "ExcelData")
% б: результат выгружается в файл result.xlsx
%---------------------------------------------------------------------------
% настройка исходной таблицы:
P_Dbl_TDC=540; %положение ВМТ такта сжатия (в градусах ПКВ)
P_Dbl_DPKVCollum_num=1; %номер колонки в которой шкала ДПКВ
P_Dbl_Cycle_Quant=50; %сколько выгружать циклов
P_Dbl_PressureCollum_num=4; %номер колонки в которой давление
P_Dbl_DPKV_Step=0.1; %шаг выгрузки данных по шкале ДПКВ
% настройка выгружаемой таблицы:
P_Dbl_TDC_Before=-180; %количество записываемых градуов на линию сжатия 
P_Dbl_TDC_After=179.9; %количество записываемых градуов на линию расширения 
%---------------------------------------------------------------------------

clear A_Mtrx_DPKV_New T_Mtrx_ExcelData A_Mtrx_DPKV_Old T_Mtrx_OutData T_Mtrx_ExcelData_flt m;

T_Mtrx_ExcelData=ExcelData;

A_Mtrx_DPKV_New=P_Dbl_TDC_Before:P_Dbl_DPKV_Step:P_Dbl_TDC_After;
A_Mtrx_DPKV_New=A_Mtrx_DPKV_New.';
A_Mtrx_DPKV_Old=ExcelData(1:end,P_Dbl_DPKVCollum_num);

i=1;
while 1
  if (round(A_Mtrx_DPKV_Old(i),1)==(P_Dbl_TDC+P_Dbl_TDC_Before))
      break;
  end
  i=i+1;
end
T_Mtrx_ExcelData(1:(i-1),:)=[];
A_Mtrx_DPKV_Old(1:(i-1))=[];

i=1;
try
   for i=1:length(A_Mtrx_DPKV_Old)
     delta=round(A_Mtrx_DPKV_Old(i+1),1)-round(A_Mtrx_DPKV_Old(i),1);
       if (round(delta,1)~=round(P_Dbl_DPKV_Step,1))
         if delta~=-719.9
           h = msgbox(['ошибка в исходной таблице, часть данных удалена, см строку:  ' num2str(i)]);
           error_i=i;
         end
       end
   end
end
T_Mtrx_ExcelData(1:(error_i-1),:)=[];
A_Mtrx_DPKV_Old(1:(error_i-1))=[];

i=1;
while 1
  if (round(A_Mtrx_DPKV_Old(i),1)==(P_Dbl_TDC+P_Dbl_TDC_Before))
      break;
  end
  i=i+1;
end
T_Mtrx_ExcelData(1:(i-1),:)=[];
A_Mtrx_DPKV_Old(1:(i-1))=[];

row=1;
for i=1:length(A_Mtrx_DPKV_Old)
  if (round(A_Mtrx_DPKV_Old(i))>=round((P_Dbl_TDC+P_Dbl_TDC_Before),1))&(round(A_Mtrx_DPKV_Old(i))<=round((P_Dbl_TDC+P_Dbl_TDC_After),1))
    T_Mtrx_ExcelData_flt(row,:)=T_Mtrx_ExcelData(i,:);
    row=row+1;
  end
end

T_Mtrx_OutData(:,1)=A_Mtrx_DPKV_New;
P_Dbl_PDKV_SektorLength=length(A_Mtrx_DPKV_New)-1;

m=0;
for i=1:P_Dbl_Cycle_Quant
   m=m+1;
   newColl=T_Mtrx_ExcelData_flt(m:(m+P_Dbl_PDKV_SektorLength),P_Dbl_PressureCollum_num);
   T_Mtrx_OutData=[T_Mtrx_OutData newColl];
   m=m+P_Dbl_PDKV_SektorLength;
end

xlswrite("result.xlsx",T_Mtrx_OutData);

