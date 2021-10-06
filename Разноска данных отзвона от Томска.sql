-- =============================================
-- Author:		Могутов
-- Create date: 31.08.2021
-- Description:	РАЗНОСКА ДАНЫХ ОБЗВОНА ДОЛЖНИКОВ, ВЫПОЛНЕННАЯ ТОМСКОМ, ИЗ ПРЕДОСТАВЛЕННОГО ИМИ ФАЙЛА EXCEL.
--				Данные появятся во вкладке "Текущая работа" лицевого счёта потребителя.

--				Путь к файлу задан в переменной @FileName, номер вносящего записи - в @Creator
--				В файле должны быть заполнены столбцы [Дата обещанного платежа], [Дата уведомления], [Лицевой счет] и [FIO](справочно). 
--				[Лицевой счет] - это LINK в dbo.SD_Sub.
--				@F_Insert присвоить = 0. Выполнить процедуру. Сверить выданное количество строк с выведенной ниже информацией - "Записей_в_файле". В идеале должно совпасть.
--				Ещё ниже выдаётся таблица, где в случае вышеуказанного несовпадения выдаётся список ошибочных записей, 
--				в которых вместо LINK в поле [Лицевой счет] стоит N_Code (Номер лицевого счёта).
--				Если всё устраивает, присвоить @F_Insert = 1 и запустить процедуру. В сообщениях проверить, что количество затронутых строк совпало с выданным при проверке.
--				Если есть перепутанные линки с номерами ЛС, то запускаем процедуру с @F_Insert = 2 (для предпросмотра - 20)
--				
-- =============================================

SET ANSI_NULLS, QUOTED_IDENTIFIER, XACT_ABORT, NOCOUNT ON

declare @C_SQL varchar(MAX) -- текст запроса данных из файла @FileName
	  , @File_Exists INT -- проверка существования файла @FileName	 
	  , @C_SQL_Chk VARCHAR(MAX) --  запрос по выборке
	  , @C_SQL_Chk1 VARCHAR(MAX) --
	  , @C_SQL_Chk2 VARCHAR(MAX) -- 
	  , @C_SQL_Chk3 VARCHAR(MAX) --
	  , @C_SQL_Chk4 VARCHAR(MAX) --
	  , @C_SQL_Chk_NoErr VARCHAR(MAX) --
	  , @C_SQL_Chk_FxdErr VARCHAR(MAX) --
	  , @C_SQL_Ins VARCHAR(MAX) -- запрос на вставку @C_SQL_Chk**
	  , @C_SQL_Ins_NoErr VARCHAR(MAX)
	  , @C_SQL_Ins_FxdErr VARCHAR(MAX)
	  , @C_SQL_Ins1 VARCHAR(MAX) -- 
	  , @C_SQL_Err VARCHAR(MAX) -- запрос на наличие ошибок
	  , @C_SQL_Err1 VARCHAR(MAX)
	  , @C_SQL_Err2 VARCHAR(MAX)

/*---------------------------------------------------------   изменить при необходимости   ----------------------------------------------------------------------*/
	  , @FileName VARCHAR(1000) = '//OMS10WS-00017/out/have_called.xlsx' -- сетевой путь к .xlsx файлу от Томска  c результатами отзвона с соответствующими заголовками
	  , @Creator INT = -718064 -- номер вносящего записи в соотвествии с LINK таблицы CS_Users
/***************************************************************************************************************************************************************/	  
	  , @F_Insert INT = 0 -- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Управление процедурой: 0 - предпросмотр / 1 - внесение строк в dbo.SD_Sub_Work 
--																														(2 - при перепутанных с ЛС линках, 20 - предпросмотр - может немного отличаться количество из-за нескольких линков у одного человека)/  
		--!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! проверить фильтр WHERE в @C_SQL_Chk4 и таблицу ошибок!
/***************************************************************************************************************************************************************/	


EXEC Master.dbo.xp_fileexist @FileName, @File_Exists OUT

IF @File_Exists = 0
	PRINT 'Нет файла ' + @FileName + ' или доступа к нему'
ELSE
	BEGIN		
------------------------------------------------ заполнение временной таблицы из файла excel ---------------------------------------
		drop table if exists #xl_table
		create table #xl_table						
		(
			FIO nvarchar(200)
		  , D_Promise nvarchar(200)
		  , D_Call DATETIME
		  , PF_LINK INT
		)

		set  @C_SQL = 		
					 'SELECT DISTINCT' 
					+'	 [FIO]'
					+'	,[Дата обещанного платежа] AS D_Promise'
					+'	,[Дата уведомления] AS D_Call'
					+'  ,[Лицевой счет] AS PF_LINK'
					+'	from '
					+'	OpenDataSource(''Microsoft.ACE.OLEDB.12.0'',''Data Source='+@FileName+';Extended Properties=EXCEL 12.0'')...[Лист1$]'
					+'  WHERE [Лицевой счет] IS NOT NULL'

		INSERT #xl_table EXEC (@C_SQL) AS LOGIN = 'sa' 
--------------------------------------------------------------------------------------------------------------------------------------
		SET @C_SQL_Err1 = 
			'select DISTINCT
				CASE 
					WHEN FIO LIKE ''%'' + pr.C_Name1 + ''%''
						THEN ''СОПОСТАВЛЕНО ПО ЛС!''					
				END AS ERRORS
			  , ss.N_Code	  AS ЛС
			  , hc.PF_LINK    AS Link_В_Списке
			  , ss.LINK		  AS Link_В_Базе			 
			  , FIO			  AS [В Списке]
			  , pr.C_Name1 +'' '' + pr.C_Name2 + '' '' + pr.C_Name3  AS [В Базе по ЛС]
			from #xl_table hc			
				JOIN (SD_Sub   ss
				JOIN CD_Ptnr pr
					ON pr.LINK = ss.F_Ptnr)
					ON 
					'

		SET @C_SQL_Err2 = 'hc.PF_LINK = ss.N_Code	
					AND FIO LIKE ''%'' + pr.C_Name1 +'' ''+ pr.C_Name2 + ''%'''
--------------------------------------------------------------------------------------------------------------------------------------

		SET @C_SQL_Ins1 = 'INSERT INTO dbo.SD_Sub_Work 
			(F_Division
			,F_SubDivision
			,F_Sub
			,D_Date
			,C_Note
			,string1
			,S_Creator
			,S_Create_Date)
			'

		SET @C_SQL_Chk1 = CONCAT('SELECT DISTINCT
			1                             AS F_Division					-- подразделение
		  , 0                             AS F_SubDivision
		  , hc.PF_LINK                    AS F_Sub					-- LINK лицевого счёта
		  , D_Call                        AS D_Date						-- дата отзвона абоненту
		  , ''Дозвон КЦ Томск по ДЗ'' + CASE
										  WHEN D_Promise IS NOT NULL THEN
											  '' ('' + D_Promise + '')''
										  ELSE
											  ''''
									  END AS C_Note						-- ответ абонента
		  , 100004220                     AS string1					-- код работы "(91) Дозвон КЦ Томска"
		  ,', @Creator,'                      AS S_Creator
		  , GETDATE()                     AS S_Create_Date				-- дата внесения записи
		  ')

		SET @C_SQL_Chk2 = ', pr.C_Name1                    AS C_Name1_в_базе
		  , FIO							  AS В_Томске
		  , ss.N_Code					  AS ЛС
		  '

		SET @C_SQL_Chk3 = 'from #xl_table hc
			JOIN(SD_Sub   ss
			JOIN CD_Ptnr pr
				ON pr.LINK = ss.F_Ptnr)
				ON ss.LINK = hc.PF_LINK
				'
		SET @C_SQL_Chk4 = '
-- where PF_LINK = 9004 or PF_LINK = 8967 --***********************************************!!!!!!!!!!!!!!!!!!!! проверка по Линку ЛС. Отключать.
		ORDER BY F_Sub
		'

		SET @C_SQL_Err = CONCAT(@C_SQL_Err1, @C_SQL_Err2)

		SET @C_SQL_Ins = CONCAT(@C_SQL_Ins1, @C_SQL_Chk1, @C_SQL_Chk3)
		SET @C_SQL_Chk = CONCAT(@C_SQL_Chk1, @C_SQL_Chk2, @C_SQL_Chk3, @C_SQL_Chk4)
		
		SET @C_SQL_Ins_NoErr = CONCAT(@C_SQL_Ins1, @C_SQL_Chk1, @C_SQL_Chk3, 'WHERE ', REPLACE(@C_SQL_Err2, '=', '!='))
		SET @C_SQL_Ins_FxdErr = CONCAT(@C_SQL_Ins1, REPLACE(@C_SQL_Chk1
												, 'hc.PF_LINK                    AS F_Sub'
												, 'ss.LINK                    AS F_Sub')
									
									, REPLACE(@C_SQL_Chk3
												, 'ss.LINK = hc.PF_LINK'
												, @C_SQL_Err2))
		SET @C_SQL_Chk_NoErr =CONCAT(@C_SQL_Chk1, @C_SQL_Chk2, @C_SQL_Chk3, 'WHERE ', REPLACE(@C_SQL_Err2, '=', '!='), @C_SQL_Chk4) --  данные без ошибок в линке
		SET @C_SQL_Chk_FxdErr =CONCAT(REPLACE(@C_SQL_Chk1
												, 'hc.PF_LINK                    AS F_Sub'
												, 'ss.LINK                    AS F_Sub')
									, @C_SQL_Chk2
									, REPLACE(@C_SQL_Chk3
												, 'ss.LINK = hc.PF_LINK'
												, @C_SQL_Err2)
									, @C_SQL_Chk4)  --  данные с исправленными в соответствии с ЛС ошибками в линке 
		
------------------------------------------------ внесение записей в базу и их просмотр ---------------------------------------
		IF @F_Insert = 1
			BEGIN
				SET NOCOUNT OFF	-- в сообщениях выводится количество строк вставки. Сравнить с контрольной выборкой в результатах.
				exec (@C_SQL_Ins)
				SET NOCOUNT ON	
				exec (@C_SQL_Chk)			
			END

		ELSE IF @F_Insert = 2
			BEGIN
				-----вставка-----
				SET NOCOUNT OFF	-- в сообщениях выводится количество строк вставки.
				--exec (@C_SQL_Ins_NoErr)
				exec (@C_SQL_Ins_FxdErr)
				-----просмотр-----
				SET NOCOUNT ON	
				exec (@C_SQL_Chk_NoErr)
				exec (@C_SQL_Chk_FxdErr)				
			END

		ELSE IF @F_Insert = 20
			BEGIN			 				
				exec (@C_SQL_Chk_NoErr)
				exec (@C_SQL_Chk_FxdErr)
				SELECT COUNT(hc.PF_LINK) AS Записей_в_файле  
					FROM #xl_table hc
			END

		ELSE
			BEGIN
				exec (@C_SQL_Chk)
				SELECT COUNT(hc.PF_LINK) AS Записей_в_файле  
					FROM #xl_table hc
	------------------------------------------------проверка наличия вместо линков лицевых счетов в исходном файле ---------------------------------------
				
				EXEC(@C_SQL_Err)
			
			END
	------------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------------
	END
 
				
				---***************** Сверка Линка ЛС и Фамилий
				--Select 
				--top 100 * 
				--	from dbo.SD_Sub ss 
				--	join CD_Ptnr pr
				--	on pr.LINK = ss.F_Ptnr
				--	where ss.LINK = 208
