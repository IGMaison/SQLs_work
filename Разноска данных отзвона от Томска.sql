-- =============================================
-- Author:		�������
-- Create date: 31.08.2021
-- Description:	�������� ����� ������� ���������, ����������� �������, �� ���������������� ��� ����� EXCEL.
--				������ �������� �� ������� "������� ������" �������� ����� �����������.

--				���� � ����� ����� � ���������� @FileName, ����� ��������� ������ - � @Creator
--				� ����� ������ ���� ��������� ������� [���� ���������� �������], [���� �����������], [������� ����] � [FIO](���������). 
--				[������� ����] - ��� LINK � dbo.SD_Sub.
--				@F_Insert ��������� = 0. ��������� ���������. ������� �������� ���������� ����� � ���������� ���� ����������� - "�������_�_�����". � ������ ������ ��������.
--				��� ���� ������� �������, ��� � ������ �������������� ������������ ������� ������ ��������� �������, 
--				� ������� ������ LINK � ���� [������� ����] ����� N_Code (����� �������� �����).
--				���� �� ����������, ��������� @F_Insert = 1 � ��������� ���������. � ���������� ���������, ��� ���������� ���������� ����� ������� � �������� ��� ��������.
--				���� ���� ������������ ����� � �������� ��, �� ��������� ��������� � @F_Insert = 2 (��� ������������� - 20)
--				
-- =============================================

SET ANSI_NULLS, QUOTED_IDENTIFIER, XACT_ABORT, NOCOUNT ON

declare @C_SQL varchar(MAX) -- ����� ������� ������ �� ����� @FileName
	  , @File_Exists INT -- �������� ������������� ����� @FileName	 
	  , @C_SQL_Chk VARCHAR(MAX) --  ������ �� �������
	  , @C_SQL_Chk1 VARCHAR(MAX) --
	  , @C_SQL_Chk2 VARCHAR(MAX) -- 
	  , @C_SQL_Chk3 VARCHAR(MAX) --
	  , @C_SQL_Chk4 VARCHAR(MAX) --
	  , @C_SQL_Chk_NoErr VARCHAR(MAX) --
	  , @C_SQL_Chk_FxdErr VARCHAR(MAX) --
	  , @C_SQL_Ins VARCHAR(MAX) -- ������ �� ������� @C_SQL_Chk**
	  , @C_SQL_Ins_NoErr VARCHAR(MAX)
	  , @C_SQL_Ins_FxdErr VARCHAR(MAX)
	  , @C_SQL_Ins1 VARCHAR(MAX) -- 
	  , @C_SQL_Err VARCHAR(MAX) -- ������ �� ������� ������
	  , @C_SQL_Err1 VARCHAR(MAX)
	  , @C_SQL_Err2 VARCHAR(MAX)

/*---------------------------------------------------------   �������� ��� �������������   ----------------------------------------------------------------------*/
	  , @FileName VARCHAR(1000) = '//OMS10WS-00017/out/have_called.xlsx' -- ������� ���� � .xlsx ����� �� ������  c ������������ ������� � ���������������� �����������
	  , @Creator INT = -718064 -- ����� ��������� ������ � ����������� � LINK ������� CS_Users
/***************************************************************************************************************************************************************/	  
	  , @F_Insert INT = 0 -- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! ���������� ����������: 0 - ������������ / 1 - �������� ����� � dbo.SD_Sub_Work 
--																														(2 - ��� ������������ � �� ������, 20 - ������������ - ����� ������� ���������� ���������� ��-�� ���������� ������ � ������ ��������)/  
		--!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! ��������� ������ WHERE � @C_SQL_Chk4 � ������� ������!
/***************************************************************************************************************************************************************/	


EXEC Master.dbo.xp_fileexist @FileName, @File_Exists OUT

IF @File_Exists = 0
	PRINT '��� ����� ' + @FileName + ' ��� ������� � ����'
ELSE
	BEGIN		
------------------------------------------------ ���������� ��������� ������� �� ����� excel ---------------------------------------
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
					+'	,[���� ���������� �������] AS D_Promise'
					+'	,[���� �����������] AS D_Call'
					+'  ,[������� ����] AS PF_LINK'
					+'	from '
					+'	OpenDataSource(''Microsoft.ACE.OLEDB.12.0'',''Data Source='+@FileName+';Extended Properties=EXCEL 12.0'')...[����1$]'
					+'  WHERE [������� ����] IS NOT NULL'

		INSERT #xl_table EXEC (@C_SQL) AS LOGIN = 'sa' 
--------------------------------------------------------------------------------------------------------------------------------------
		SET @C_SQL_Err1 = 
			'select DISTINCT
				CASE 
					WHEN FIO LIKE ''%'' + pr.C_Name1 + ''%''
						THEN ''������������ �� ��!''					
				END AS ERRORS
			  , ss.N_Code	  AS ��
			  , hc.PF_LINK    AS Link_�_������
			  , ss.LINK		  AS Link_�_����			 
			  , FIO			  AS [� ������]
			  , pr.C_Name1 +'' '' + pr.C_Name2 + '' '' + pr.C_Name3  AS [� ���� �� ��]
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
			1                             AS F_Division					-- �������������
		  , 0                             AS F_SubDivision
		  , hc.PF_LINK                    AS F_Sub					-- LINK �������� �����
		  , D_Call                        AS D_Date						-- ���� ������� ��������
		  , ''������ �� ����� �� ��'' + CASE
										  WHEN D_Promise IS NOT NULL THEN
											  '' ('' + D_Promise + '')''
										  ELSE
											  ''''
									  END AS C_Note						-- ����� ��������
		  , 100004220                     AS string1					-- ��� ������ "(91) ������ �� ������"
		  ,', @Creator,'                      AS S_Creator
		  , GETDATE()                     AS S_Create_Date				-- ���� �������� ������
		  ')

		SET @C_SQL_Chk2 = ', pr.C_Name1                    AS C_Name1_�_����
		  , FIO							  AS �_������
		  , ss.N_Code					  AS ��
		  '

		SET @C_SQL_Chk3 = 'from #xl_table hc
			JOIN(SD_Sub   ss
			JOIN CD_Ptnr pr
				ON pr.LINK = ss.F_Ptnr)
				ON ss.LINK = hc.PF_LINK
				'
		SET @C_SQL_Chk4 = '
-- where PF_LINK = 9004 or PF_LINK = 8967 --***********************************************!!!!!!!!!!!!!!!!!!!! �������� �� ����� ��. ���������.
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
		SET @C_SQL_Chk_NoErr =CONCAT(@C_SQL_Chk1, @C_SQL_Chk2, @C_SQL_Chk3, 'WHERE ', REPLACE(@C_SQL_Err2, '=', '!='), @C_SQL_Chk4) --  ������ ��� ������ � �����
		SET @C_SQL_Chk_FxdErr =CONCAT(REPLACE(@C_SQL_Chk1
												, 'hc.PF_LINK                    AS F_Sub'
												, 'ss.LINK                    AS F_Sub')
									, @C_SQL_Chk2
									, REPLACE(@C_SQL_Chk3
												, 'ss.LINK = hc.PF_LINK'
												, @C_SQL_Err2)
									, @C_SQL_Chk4)  --  ������ � ������������� � ������������ � �� �������� � ����� 
		
------------------------------------------------ �������� ������� � ���� � �� �������� ---------------------------------------
		IF @F_Insert = 1
			BEGIN
				SET NOCOUNT OFF	-- � ���������� ��������� ���������� ����� �������. �������� � ����������� �������� � �����������.
				exec (@C_SQL_Ins)
				SET NOCOUNT ON	
				exec (@C_SQL_Chk)			
			END

		ELSE IF @F_Insert = 2
			BEGIN
				-----�������-----
				SET NOCOUNT OFF	-- � ���������� ��������� ���������� ����� �������.
				--exec (@C_SQL_Ins_NoErr)
				exec (@C_SQL_Ins_FxdErr)
				-----��������-----
				SET NOCOUNT ON	
				exec (@C_SQL_Chk_NoErr)
				exec (@C_SQL_Chk_FxdErr)				
			END

		ELSE IF @F_Insert = 20
			BEGIN			 				
				exec (@C_SQL_Chk_NoErr)
				exec (@C_SQL_Chk_FxdErr)
				SELECT COUNT(hc.PF_LINK) AS �������_�_�����  
					FROM #xl_table hc
			END

		ELSE
			BEGIN
				exec (@C_SQL_Chk)
				SELECT COUNT(hc.PF_LINK) AS �������_�_�����  
					FROM #xl_table hc
	------------------------------------------------�������� ������� ������ ������ ������� ������ � �������� ����� ---------------------------------------
				
				EXEC(@C_SQL_Err)
			
			END
	------------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------------
	END
 
				
				---***************** ������ ����� �� � �������
				--Select 
				--top 100 * 
				--	from dbo.SD_Sub ss 
				--	join CD_Ptnr pr
				--	on pr.LINK = ss.F_Ptnr
				--	where ss.LINK = 208
