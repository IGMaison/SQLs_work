-- ЧЕРНОВИК


declare @C_SQL varchar(MAX)
declare @t table (n nvarchar(100) )


set  @C_SQL = 		''
			--'INSERT   -- выборка в архив юл
			-- [CUS].[CallArch_EE]
			-- SELECT DISTINCT
			--	inf.SS_LINK
			--	,inf.TRYRESULT
			--	,inf.D_ZVON
			--	,inf.HOUR
			--	,inf.MIN
			--	,inf.TELEFON
			--	,inf.N_DZ
			--	,inf.N_SHTRAF
			--	,inf.S_DOLG				
			--	,inf.PENI
			--	,inf.D_DOLG			
			--FROM [DBF_INFIN_PF_TST2]...[Obzvon1] inf 
			----JOIN [CUS].[CallArch_EE] omn 
			----ON inf.SS_LINK = omn.F_SS_LINK
			--WHERE inf.TRYRESULT = 10
			--	OR inf.NOM_ZVON = 3
			--	'

			--'DELETE   -- удаление обзвоненных юл			 		
			-- [DBF_INFIN_PF_TST2]...[Obzvon1] 
			--WHERE TRYRESULT = 10
			--	OR NOM_ZVON = 3
			--	'

		


			--	'select top 20 * from [DBF_INFIN_PF_TST2]...[Obzvon1]'
			--  'select top 10 * from OpenDataSource(''Microsoft.ACE.OLEDB.12.0'',''Data Source=\\infinity\CallCenter;Extended Properties=DBASE IV'')...obzvon1'
			-- 'SELECT TOP 20 TELEFON FROM OPENROWSET (''MICROSOFT.ACE.OLEDB.12.0'',''dBASE IV;CharacterSet=866;DATABASE=\\infinity\CallCenter\2017'',''Select * from obzvtst.DBF'')'

		EXECUTE ( @C_SQL )  AS LOGIN = 'sa'

	
SELECT TOP (
				select 
					2 * avg(cnt.N_Calls_Cnt) 
				FROM (
						select DISTINCT TOP(10) 
							(COUNT(ph.C_Telefon) OVER(PARTITION BY D_Zvon ORDER BY D_Zvon DESC)) N_Calls_Cnt 
						FROM [CUS].[CallArch_EE] ph
				) cnt
		) --  Удвоенное Среднее количество звонков за 10 дней.
	--@Session                                  AS Session_Id
	  debts.[N_Code]                            AS N_LS
	, debts.[C_Number]                          AS [Doc_Num]
	, debts.[C_FIO]                             AS [Consumer]
	, debts.[C_Analyst]
	, debts.[N_SummaEE]						    AS N_DZ
	, debts.[N_Shtraf]						    AS N_SHTRAF
	--, t.TEL									 AS TELEFON
	, debts.[N_MainRealiz]						AS S_DOLG
	, debts.[N_Peni]                            AS PENI
	, debts.[C_Call_Back_Tel]                   AS TEL_D
	, DATEADD(d, -2, GETDATE()) --[D_DateCmd]					             AS D_DOLG
	, ROW_NUMBER() OVER (ORDER BY debts.N_SummaEE DESC) AS [ID]  --20210823
	, 0                                         AS REZ_ZVON
	, 0                                         AS [MIN]
	, 0                                         AS [HOUR]
	, 0                                         AS TRYRESULT
	, 0                                         AS NOM_ZVON
	, '20000101'--@DateConst                                AS D_ZVON
	, debts.F_Subscr									 AS SS_LINK  --20210908
   

	,old.D_Zvon
	,DATEDIFF(D, old.D_Zvon, GETDATE()) --

FROM [CUS].[RPT_406_Debts_Inf_New4_Email_EE] debts
LEFT JOIN 
	(
		SELECT
		  arh.F_SS_LINK
		, MAX(arh.D_Zvon) D_Zvon
		FROM [CUS].[CallArch_EE] arh
		WHERE DATEDIFF(D, arh.D_Zvon, GETDATE()) < 12-- нет в архиве ранее @N_дней. Даты!!
		GROUP BY arh.F_SS_LINK
	
	) old
	
	ON debts.F_Subscr = old.F_SS_LINK
	
				
WHERE debts.N_SummaEE >= 100 --@MinSumm	
	AND old.F_SS_LINK is null -- нет в архиве ранее @N_дней. Даты!!
--ORDER BY 
--ORDER BY [ID] ASC

--select F_SS_LINK, MAX(D_Zvon) FROM [CUS].[CallArch_EE] WHERE DATEDIFF(D, D_Zvon, GETDATE()) > 12 GROUP BY F_SS_LINK 
--select 2 * avg(N_Calls_Cnt) FROM (select DISTINCT TOP(10) (COUNT(C_Telefon) OVER(PARTITION BY D_Zvon ORDER BY D_Zvon DESC)) N_Calls_Cnt FROM [CUS].[CallArch_EE]) cnt --  Удвоенное Среднее количество звонков за 10 дней.







		--EXEC('UPDATE a 
		--		SET a.SS_LINK = Max(b.LINK) OVER (PARTITION BY  a.SS_LINK)
		--		FROM
		--			OpenDataSource(''Microsoft.ACE.OLEDB.12.0'',''Data Source=\\infinity\Users\test\;Extended Properties=DBASE IV'')...obzvon1 a 
		--			join SD_Subscr b 
		--			ON a.N_LS = b.N_Code') AS LOGIN = 'sa'


--'SELECT   --- ВСТАВКА ЛИНКОВ
--				a.SS_LINK
				
						 
--				FROM
--					[DBF_INFIN_PF_TST2]...[Obzvon1] a 
--					join (select max(LINK) LINK, N_Code from SD_Subscr group by N_Code) b 
--					ON a.N_LS = b.N_Code'

 --INSERT @t EXECUTE ( @C_SQL )  AS LOGIN = 'sa'

 
