%---------------------------------------------------------------------------
P_Dbl_TDC=540; %положение ВМТ такта сжатия (в градусах ПКВ)
P_Dbl_DPKVCollum_num=1; %номер колонки в которой шкала ДПКВ
P_Dbl_Cycle_Quant=50; %сколько выгружать циклов

P_Dbl_TDC_Before=-140; %количество записываемых градуов на линию сжатия 
P_Dbl_TDC_After=100; %количество записываемых градуов на линию расширения 

P_Dbl_PressureCollum_num=4; %номер колонки в которой давление
P_Dbl_DPKV_Step=0.1; %шаг выгрузки данных по шкале ДПКВ
%---------------------------------------------------------------------------

clear A_Mtrx_DPKV_New T_Mtrx_ExcelData A_Mtrx_DPKV_Old T_Mtrx_OutData T_Mtrx_ExcelData_flt m;

T_Mtrx_ExcelData=ExcelData;

A_Mtrx_DPKV_New=P_Dbl_TDC_Before:P_Dbl_DPKV_Step:P_Dbl_TDC_After;
A_Mtrx_DPKV_New=A_Mtrx_DPKV_New.';
A_Mtrx_DPKV_Old=ExcelData(1:end,P_Dbl_DPKVCollum_num);

row=1;
for i=1:length(A_Mtrx_DPKV_Old)
  if (round(A_Mtrx_DPKV_Old(i),1)>=(P_Dbl_TDC+P_Dbl_TDC_Before))&(round(A_Mtrx_DPKV_Old(i),1)<=(P_Dbl_TDC+P_Dbl_TDC_After))
    T_Mtrx_ExcelData_flt(row,:)=T_Mtrx_ExcelData(i,:);
    row=row+1;
  end
end

T_Mtrx_OutData(:,1)=A_Mtrx_DPKV_New;
P_Dbl_PDKV_SektorLength=((P_Dbl_TDC_After-P_Dbl_TDC_Before)/P_Dbl_DPKV_Step);
m=1;
for i=1:P_Dbl_Cycle_Quant
   newColl=T_Mtrx_ExcelData_flt(m:(m+P_Dbl_PDKV_SektorLength),P_Dbl_PressureCollum_num);
   T_Mtrx_OutData=[T_Mtrx_OutData newColl];
   m=m+P_Dbl_PDKV_SektorLength+1;
end

xlswrite("result.xlsx",T_Mtrx_OutData);

