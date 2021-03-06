SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================

-- Description:	Отправка сообщений по емэйл ФЛ, у кого есть счётчики тепла и емэйл, на основе процедуры [*].[*] 100 Журнал всех ЛС

-- =============================================
ALTER PROCEDURE [*].[*]

	@N_Period INT,
	@F_Sub INT,
	@F_Division INT = 1,
	@C_Personal_Account VARCHAR(3)
AS
BEGIN
	SET NOCOUNT, XACT_ABORT, ANSI_PADDING, ANSI_WARNINGS, ARITHABORT, CONCAT_NULL_YIELDS_NULL ON
	SET NUMERIC_ROUNDABORT, CURSOR_CLOSE_ON_COMMIT OFF
	SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED
	
	DECLARE @D_Date              SMALLDATETIME,	       
	        @D_Date2             SMALLDATETIME
	
	SET @D_Date = GETDATE()	
	SET @D_Date2 = CASE 
	                    WHEN YEAR(@D_Date) * 100 + MONTH(@D_Date) = @N_Period THEN @D_Date
	                    ELSE dbo.CF_Month_Date_End_N_Period(@N_Period)
	               END

	
	DECLARE @F_Personal_Account INT = NULL
	
	SELECT @F_Personal_Account = spar.LINK 
	FROM   SD_Personal_Account spar
	WHERE  spar.F_Division = @F_Division
	       AND spar.string2 IS NOT NULL
	       AND CONCAT(LEFT(spar.C_Name, 1), RIGHT(spar.C_Name, 2)) = @C_Personal_Account

	
	DECLARE @SIT_Heat     INT,
	        @SIT_GVS      INT
	
	SELECT @SIT_Heat = fsi.LINK
	FROM   FS_Sale_It AS fsi
	WHERE  fsi.C_Const = 'SIT_Heat'
	
	SELECT @SIT_GVS = fsi.LINK
	FROM   FS_Sale_It AS fsi
	WHERE  fsi.C_Const = 'SIT_GVS'
	

	----------------------------------------------------------------------------------
	-- Наличие приборов
	IF OBJECT_ID('tempdb..#Devs') IS NOT NULL
	    DROP TABLE #Devs 
	CREATE TABLE #Devs
	(
		F_Sub          INT,
		F_Division        INT,
		B_Heat_Dev     BIT,
		B_HW_Dev       BIT
		PRIMARY KEY(F_Division, F_Sub)
	)
	INSERT INTO #Devs
	  (
	    F_Sub,
	    F_Division,
	    B_Heat_Dev,
	    B_HW_Dev
	  )

	SELECT --top 10000   --*******************************************************************************************
		
		ss.LINK              AS F_Sub,
	        ss.F_Division,
	        MAX(CASE WHEN erp.F_Sale_It = @SIT_Heat THEN 1 ELSE 0 END) 
	        B_Heat_Dev,
	        MAX(CASE WHEN erp.F_Sale_It = @SIT_GVS THEN 1 ELSE 0 END) AS B_HW_Dev
	FROM   
	       SD_Sub AS ss
	       JOIN SD_Personal_Account AS spar
	            ON  spar.LINK = ss.F_Personal_Account
	            AND spar.F_Division = ss.F_Division
	       JOIN ED_Reg_Pts  AS erp
	            ON  erp.F_Sub = ss.LINK
	            AND erp.F_Division = ss.F_Division
	            AND erp.D_Date_Begin <= @D_Date2
	            AND ISNULL(erp.D_Date_End, '20790606') > @D_Date2 
	                --AND (erp.D_Date_End > @D_Date2 OR erp.D_Date_End IS NULL)
	       JOIN ED_Devs_Pts  AS edp
	            ON  edp.F_Reg_Pts = erp.LINK
	            AND edp.F_Division = erp.F_Division
	       JOIN ED_Devs      AS ed
	            ON  ed.LINK = edp.F_Devs
	            AND ed.F_Division = edp.F_Division
	            AND ed.D_Setup_Date <= @D_Date2
	            AND ISNULL(ed.D_Replace_Date, '20790606') > @D_Date2
	WHERE  ss.F_Division = @F_Division
	       AND ss.B_EE = 0
	       AND (ss.LINK = @F_Sub OR @F_Sub IS NULL)
	       AND (
	               spar.LINK = @F_Personal_Account
	               OR @F_Personal_Account IS NULL
	           )
	GROUP BY
	       ss.LINK,
	       ss.F_Division
	
		--select count(F_Sub) from #Devs	

---======================================================================================================
	
	
	
 --==============================================
 -- Получаем список адресов
 --==============================================


 IF OBJECT_ID('tempdb..#Email') IS NOT NULL
     DROP TABLE #Email




SELECT DISTINCT

	LEFT(spar.C_Name, 1), RIGHT(spar.C_Name, 2))       AS [Округ],
	spar.String2					   AS [Адрес],
	spar.String3					   AS [email],   
        CASE 
            WHEN    ss.B_Inactive = 1 
		 OR ss.D_Date_End < @D_Date 

	    	THEN 'Закрыт'
            	ELSE sss2.C_Name
            END                          		   AS [Статус],
       
        cp.C_Email                 			   AS [Адрес электронной почты],
        ISNULL(dev.B_Heat_Dev, 0)  			   AS [ИПУ Отоп]
      
INTO   #Email 
FROM   SD_Sub AS ss
       LEFT JOIN SS_Sub_St AS sss2
            ON  sss2.LINK = ss.F_Sub_St
       LEFT JOIN #Devs  AS dev
            ON  dev.F_Sub = ss.LINK
            AND dev.F_Division = ss.F_Division
       JOIN SD_Personal_Account AS spar
            ON  spar.F_Division = ss.F_Division
            AND spar.LINK = ss.F_Personal_Account      
       JOIN CD_Prtn AS cp 
            ON  cp.LINK = ss.F_Prtn
            AND cp.F_Division = ss.F_Division     
     
     
WHERE   (ss.LINK = @F_Sub OR @F_Sub IS NULL)
        AND ss.B_EE = 0
        AND ss.F_Division = @F_Division
        AND (
               spar.LINK = @F_Personal_Account
               OR @F_Personal_Account IS NULL
           )
	AND ISNULL(dev.B_Heat_Dev, 0) = 1
	AND cp.C_Email LIKE '%@%'
	AND (CASE 
       		 WHEN      ss.B_Inactive = 1
			OR ss.D_Date_End < @D_Date 

		 THEN 'Закрыт'
            	 ELSE sss2.C_Name

	     END) != 'Закрыт'          
GROUP BY
       CONCAT(
	   LEFT(spar.C_Name, 1), 
	   RIGHT(spar.C_Name, 2)),
	   spar.String2,	
	   spar.String3,	
       CASE 
            WHEN   ss.B_Inactive = 1 
		OR ss.D_Date_End < @D_Date THEN 'Закрыт'
            ELSE sss2.C_Name
       END,       
       ss.C_Number,
       cp.C_Email,
       ISNULL(dev.B_Heat_Dev, 0)


--select  * from #Email


----==============================================
-- -- Отправка
-- --==============================================

DECLARE @EMAIL VARCHAR (MAX)
		,@district VARCHAR (100)
		,@adress VARCHAR (100)
		,@call_back_email VARCHAR (50)
		,@flag int = 0
		,@C_Body VARCHAR(MAX)
		,@C_Subject VARCHAR (200) = 'Уведомление о предоставлении информации'
		,@C_Body_format VARCHAR (20) = 'HTML'
		,@err INT = 0


DECLARE @CURSOR CURSOR

SET @CURSOR  = CURSOR SCROLL
FOR
	SELECT [Округ], [Адрес электронной почты], [Адрес], email FROM #Email AS e

OPEN @CURSOR

FETCH NEXT FROM @CURSOR INTO @district, @EMAIL, @adress, @call_back_email


WHILE (@@FETCH_St = 0)-- and @flag < 1)--ограничение количества писем для тестов!!!!!!!!!!!!!!!!!!!**********************************************
BEGIN

BEGIN TRY
	SET @flag = @flag + 1 

	

	SET @EMAIL = [dbo].[OF_GetVarMailErr](@EMAIL, getdate(),'u', 'E-Mail - исправить')-- функция исправления емэйла
	SET @C_Body = CONCAT('<html><head><meta http-equiv=Content-Type content="text/html; charset=1251"><meta name=Generator content="Microsoft Word 15 (filtered medium)"></head><body lang=RU link="#0563C1" vlink="#954F72">*********'
					,@district
					,' ******, расположенный по адресу: '
					,@adress
					,', посредством специально оборудованного ящика для приема корреспонденции или специалисту в окне*********на электронную почту абонентского отдела '
					,@district
					,': </span><a href="mailto:'
					,@call_back_email
					,'"><span style=''font-size:15.0pt;font-family:"Times New Roman",serif''>'
					,@call_back_email
					,'</span></a>*********** необходимо предоставить соответствующие документы в абонентский отдел.</span></p></div></body></html>')

 

	EXEC msdb.dbo.sp_send
					@profile_name		=	'Informir'										
					,@recipients		=	@EMAIL 									
					,@subject		=	@C_Subject
					,@body			=	@C_Body
					,@body_format		=	@C_Body_format
					,@file_att ='D:\temp\logoNew23092016.png';	---Файл во вложении



END TRY
BEGIN CATCH
	SELECT @@ERROR, CONCAT('Использовано ', @flag, ' адреса(-ов)')
	SET @err = @err + 1
	SET @flag = @flag - 1
END CATCH

/*Выбираем следующую строку*/
FETCH NEXT FROM @CURSOR INTO @district, @EMAIL, @adress, @call_back_email
END

CLOSE @CURSOR
DEALLOCATE @CURSOR


-- контрольное письмо---------------------------------
SELECT @adress = CONCAT('<font color=''red''>****Отправлено по ' , @flag , ' адресу(-ам), ошибок - ', @err,' ****</font>', @adress)
SET @C_Body = CONCAT('**********')

 
EXEC msdb.dbo.sp_send
					@profile_name		=	'Informir'												
					,@recipients		=	'c----'
					,@copy_recipients	=	'mo----'					
					,@subject		=	@C_Subject
					,@body			=	@C_Body
					,@body_format		=	@C_Body_format
					,@file_att ='D:\temp\logoNew23092016.png';	
-------------------------------------

PRINT(CONCAT('Использовано ', @flag, ' адреса(-ов).', 'Ошибок - ', @err))
--select  * from #Email

END