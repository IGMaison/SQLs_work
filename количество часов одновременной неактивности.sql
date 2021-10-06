SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Могутов
-- Create date: 2021-09-20
-- Description:	Возвращает количество часов одновременной неактивности (N_Rate = 0) по подразделениям 1 и 4 для конкретного счётчика в заданных временнЫх рамках.
-- =============================================
ALTER FUNCTION [OF].[Inactivity_RegPTS_Time_HH]
(
		  @D_Date					SMALLDATETIME 	--  Дата показания счётчика 
		, @D_Date_Prev					SMALLDATETIME 	--  Дата предыдущего показания счётчика		
		, @C_Serial_Number				VARCHAR(100)		
)
RETURNS INT
AS
BEGIN


	DECLARE	  			
			  @F_Reg_PTS_Div1				INT -- линк на уч.показатель прибора по подразд. 1
			, @F_Reg_PTS_Div4				INT -- линк на уч.показатель прибора по подразд. 4			
			, @Count_Div1					INT -- подсчёт количества переключений в пределах диапазона текущего и предыдущего показаний счётчика по подразделению 1
			, @Count_Div4					INT -- подсчёт количества переключений в пределах диапазона текущего и предыдущего показаний счётчика по подразделению 4
			, @Return_Inactivty_Hours			INT -- выходное значение часов неактивности по обоим направлениям.
			

	DECLARE --таблица с линками на учётный показатель для данного счётчика по подразделению 1 и 4 (они разные для одного и того же счётчика + он представлен в таблице счётчиков дважды!)
             @Reg_Pts_Links TABLE (
                               [F_Division]     INT  NULL,
                               [PTS_Link]		INT  NULL
										)

            


	
	INSERT INTO  @Reg_Pts_Links
	SELECT DISTINCT 
		  d.F_Division
		, rp.LINK PTS_Link
		FROM ED_DEV d
			JOIN ED_DEV_Pts dp 
				ON dp.F_DEV = d.LINK
			JOIN ED_Reg_Pts rp 
				ON rp.LINK = dp.F_Reg_Pts
			AND C_Serial_Number = @C_Serial_Number

	SELECT 
		@F_Reg_PTS_Div1 = PTS_Link 
		FROM @Reg_Pts_Links 
		WHERE F_Division = 1
	SELECT 
		@F_Reg_PTS_Div4 = PTS_Link 
		FROM @Reg_Pts_Links
		WHERE F_Division = 4

	 SELECT DISTINCT @Count_Div1 = count(LINK) FROM dbo.ED_Reg_Pts_Act
	 WHERE
		 D_Date > @D_Date_Prev
		AND D_Date < @D_Date
		AND F_Reg_Pts = @F_Reg_PTS_Div1

	 SELECT DISTINCT @Count_Div4 = count(LINK) FROM dbo.ED_Reg_Pts_Act
	 WHERE
		 D_Date > @D_Date_Prev
		AND D_Date < @D_Date
		AND F_Reg_Pts = @F_Reg_PTS_Div4


	SELECT @Return_Inactivty_Hours = SUM(DATEDIFF(hh, Activity_Begin, Activity_End))
	FROM
		(
			SELECT DISTINCT 

				Q_Div1.D_Date_Case_Device D_Date_Case_Device1
				, Q_Div1.D_Date_End D_Date_End1
				, Q_Div4.D_Date_Case_Device D_Date_Case_Device4
				, Q_Div4.D_Date_End D_Date_End4
				, CASE WHEN Q_Div1.D_Date_Case_Device > Q_Div4.D_Date_Case_Device THEN 
				Q_Div1.D_Date_Case_Device ELSE 
				Q_Div4.D_Date_Case_Device END Activity_Begin
				, CASE WHEN Q_Div1.D_Date_End < Q_Div4.D_Date_End THEN 
				Q_Div1.D_Date_End ELSE 
				Q_Div4.D_Date_End END Activity_end
				, Q_Div1.Activity Activity_Dev1
				, Q_Div4.Activity Activity_Dev4
				, CASE WHEN Q_Div1.Activity = 1 or Q_Div4.Activity = 1
					 THEN 1 
					 ELSE 0 
				  END Real_Activity
			FROM 
					(SELECT top (@Count_Div1 + 1)
						D_Date
						,CASE WHEN D_Date < @D_Date_Prev THEN
						@D_Date_Prev ELSE
						D_Date end D_Date_Case_Device
						, LEAD(D_Date, 1, @D_Date)  OVER (ORDER BY D_Date) D_Date_End
						, N_Rate Activity
					FROM dbo.ED_Reg_Pts_Act
					WHERE F_Reg_Pts = @F_Reg_PTS_Div1    
					AND D_Date < @D_Date
					ORDER BY D_Date DESC)
					Q_Div1

				join (  SELECT TOP (@Count_Div4 + 1)
							D_Date
							,CASE WHEN D_Date < @D_Date_Prev THEN
							@D_Date_Prev ELSE
							D_Date end D_Date_Case_Device
							, LEAD(D_Date, 1, @D_Date)  OVER (ORDER BY D_Date) D_Date_End
							, N_Rate Activity
						FROM dbo.ED_Reg_Pts_Act
						WHERE  
							F_Reg_Pts = @F_Reg_PTS_Div4
							AND D_Date < @D_Date
						ORDER BY 
							D_Date DESC
					) Q_Div4

					ON  Q_Div4.D_Date >= Q_Div1.D_Date 
						AND Q_Div4.D_Date < Q_Div1.D_Date_End
						OR
						Q_Div1.D_Date >= Q_Div4.D_Date 
						AND Q_Div1.D_Date < Q_Div4.D_Date_End

		) Q_Div1_Div4
	GROUP BY 
		Q_Div1_Div4.Real_Activity
	HAVING 
		Q_Div1_Div4.Real_Activity = 0
	
	
	RETURN CASE
				WHEN @Return_Inactivty_Hours IS NULL
					THEN 0
				ELSE @Return_Inactivty_Hours
			END
END
GO

