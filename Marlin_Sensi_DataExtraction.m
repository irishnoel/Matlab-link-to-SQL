

%% MS SQL: Initialize MS SQL database 
clear
format long
cn_str = 'Provider=sqloledb;Data Source=SWDPDEDCML101; Initial Catalog=Reporting_DB_HRE;  Integrated Security=SSPI';
DB = adodb_connect(cn_str, 240);

%% MS SQL: Run statement with multiple queries -- query we use is loaded from file
% sql = 'Select MAX(TimeID) FROM Star.Fact_Performance';


%fid = fopen('C:\Matlab_code\ADO SQL Conections\New_NewYork PnL Query_ver_9_finished.txt');
sqlcommand = fileread('\\depfa.loc\dfs_shares\001890_DP_USNY_RISK_IT_Market_Risk\Reports\DailyReport\Local Sensis and VaR\DAILY RUN KRM\New_PnL_files\SQL_Queries\NewYork_SanSabia_NoTenor_Sensi.txt');
sql = sqlcommand;
[Struct, Table] = adodb_query(DB, sql);

%disp('Output in in struct format:')
%disp(Struct)
%disp('Output in in table format:')
%disp(Table)

%% Output the data

%% NY & SanSabia Data First
report_date = Table{1,1};
ddate = int2str(report_date);
fields = transpose(repmat(fieldnames(Struct), numel(Struct), 1));
datatowrite = vertcat(fields, Table);
filenameuse = ['\\depfa.loc\dfs_shares\001890_DP_USNY_RISK_IT_Market_Risk\Reports\DailyReport\Local Sensis and VaR\DAILY RUN KRM\New_PnL_files\Raw_Input_Data\Sensi_Data\NewYork\NY_Branch_Sensi_Data_' ddate  '.csv'];
cell2csv(filenameuse, datatowrite)






%% Now NY Data with Tenor Details
sqlcommand_NYTenor = fileread('\\depfa.loc\dfs_shares\001890_DP_USNY_RISK_IT_Market_Risk\Reports\DailyReport\Local Sensis and VaR\DAILY RUN KRM\New_PnL_files\SQL_Queries\TenorTest.txt');
sql = sqlcommand_NYTenor;
[Struct, Table] = adodb_query(DB, sql);


ddate = int2str(report_date);
fields = transpose(repmat(fieldnames(Struct), numel(Struct), 1));
datatowrite = vertcat(fields, Table);
filenameuse = ['\\depfa.loc\dfs_shares\001890_DP_USNY_RISK_IT_Market_Risk\Reports\DailyReport\Local Sensis and VaR\DAILY RUN KRM\New_PnL_files\Raw_Input_Data\Sensi_Data\NewYork\NY_Branch_Tenor_SensiData_' ddate  '.csv'];
cell2csv(filenameuse, datatowrite)



DB.release;
