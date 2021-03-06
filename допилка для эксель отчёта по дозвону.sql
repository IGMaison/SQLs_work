
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		
-- Create date: 20161021
-- ALTER DATE:	20210928 Могутов. Отображение данных теперь формируется прямо с таблицы автодозвона (только отзвоненные) и оплаты учитываются с даты звонка.
-- Description:	Автодозвон по ФЛ
-- =============================================
ALTER procedure [*].[*]
	@N_Period int,
	@F_Division tinyint = 1,
	@N_Days INT = 0
AS
begin

	
	set nocount, xact_abort, ansi_padding, ansi_warnings, arithabort, concat_null_yields_null on
	set numeric_roundabort, cursor_close_on_commit off

	if @N_Period is null
		select @N_Period = sd.N_Year * 100 + sd.N_Month from SD_Divisions as sd where sd.LINK = @F_Division

    declare @N_Period_Prev INT
    set @N_Period_Prev = dbo.CF_Period_Prev_Next(@N_Period, 1)
  
    declare @FSC_PE_Base_Cons int
    select @FSC_PE_Base_Cons = fsc.LINK from FS_Sale as fsc where fsc.C_Const = 'FSC_PE_Base_Cons'

	if object_id('tempdb..#Calls') is not null drop table #Calls
	create table #Calls (

		F_Sub bigint,
		N_Period int,
		D_Date smalldatetime,
		C_Telephone varchar(50),
		N_Dolg money,
		C_Nam1 varchar(300),
		C_Nam2 varchar(300),
		C_Nam3 varchar(50)
		primary key clustered (F_Sub, N_Period) 
	)

declare @C_SQL varchar(MAX)

set  @C_SQL = 	'

	insert into #Calls (

		F_Sub,
		N_Period,
		D_Date,
		C_Telephone,
		N_Dolg,
		C_Nam1,
		C_Nam2,
		C_Nam3
	)
	select 
		ps.SS_LINK F_Sub,
		YEAR(ps.D_ZVON) * 100 + MONTH(ps.D_ZVON) N_Period,
		ps.D_zvon D_Date,
		ps.Telefon C_Telephone,
		ps.S_Dolg N_Dolg,
		ps.FAMIL,
		ps.IMYA,
		ps.OTCH
	from
		 OpenDataSource(''Microsoft.ACE.OLEDB.12.0'',''Data Source=\\inf\CallCenter;Extended Properties=DBASE IV'')...obzvon as ps 		
	where YEAR(ps.D_ZVON) * 100 + MONTH(ps.D_ZVON) = @N_Period
    OPTION(RECOMPILE)
'
EXECUTE ( @C_SQL ) 

    select

	poc.N_Period,
        ss.LINK,
	ss.C_Number ЖЭУ, 
	ss.N_Code ЛС,
        pt.C_Nam1 Ф, 
	pt.C_Nam2 И, 
	pt.C_Nam3 О,
        spar.C_Nam АО,
        scp.C_Address_Short [Адрес], 
	scps.C_Premise_Number [Квартира],
        poc.C_Telephone Телефон, 
	poc.N_Dolg Долг,
        poc.D_Date [Дата звонка],
	sum(fd.N_Amount) [Оплаты с момента обзвона]
    from

	#Calls poc
	join dbo.SD_Sub ss 
		on ss.LINK = poc.F_Sub 
		and ss.F_Division = @F_Division
	left join SD_Pers_Acc_Reg as spar 
		on spar.LINK = ss.F_Pers_Acc_Reg
	join SD_Conn_Pnts as scp 
		on scp.LINK = ss.F_Conn_Pnts 
		and scp.F_Division = ss.F_Division
	join SD_Conn_Pnts_Sub as scps 
		on scps.LINK = ss.F_Conn_Pnts_Sub 
		and scps.F_Division = ss.F_Division 
		and scps.F_Conn_Pnts = scp.LINK
	left join pe.FD_Pments fd 
		on fd.F_Division = ss.F_Division 
            	and fd.F_Sub = ss.LINK
            	and fd.D_Date >= poc.D_Date 
		AND case 
			when @N_Days != 0 
			then DATEADD(dd,@N_Days,poc.D_Date) 
			ELSE '20790606' 
		    END 
			>= fd.D_Date
            	and fd.F_Sale = @FSC_PE_Base_Cons
	JOIN CD_Ptnrs pt 
		ON pt.LINK = ss.F_Ptnrs
    group by

	poc.N_Period,
        ss.LINK, 
	ss.C_Number, 
	ss.N_Code,
        spar.C_Nam,
        scp.C_Address_Short, 
	scps.C_Premise_Number,
        pt.C_Nam1, 
	pt.C_Nam2, 
	pt.C_Nam3, 
	poc.C_Telephone, 
	poc.N_DOLG,
        poc.D_Date
	order by
	poc.D_Date

    OPTION(RECOMPILE)

end