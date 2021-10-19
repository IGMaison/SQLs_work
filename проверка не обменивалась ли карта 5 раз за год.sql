--проверка не обменивалась ли карта покупателя 5 раз за год

drop  table #t

create table #t (i int, a int, d int) -- старый номер, новый номер "дата" замены

insert into #t values(444, 777, 305)
insert into #t values(777, 123, 307)

insert into #t values(111, 555, 109)
insert into #t values(222, 223, 210)
insert into #t values(333, 334, 311)
insert into #t values(444, 222, 412)
insert into #t values(555, 666, 512)
insert into #t values(666, 777, 613)
insert into #t values(777, 888, 714)
insert into #t values(888, 000, 815)
insert into #t values(999, 333, 916)
insert into #t values(223, 111, 1016)
insert into #t values(456, 888, 1018)



declare
  @curr_cart_num int = 0
  ,@limit int = 5
  ,@curr_change_date int = 10000
  , @count int = 0


while @count < @limit

  begin
	if exists(
					select top 1
						old_card
					from 
						#t
					where
						new_card = @curr_cart_num
						and dt < @curr_change_date
						and dt > dateadd(yy, getdate(), -1)
					order by
						dt desc
				)
	
		begin
			select top 1			
				@curr_change_date = dt
				,@curr_cart_num = old_card
			from 
				#t
			where
				new_card = @curr_cart_num
				and dt < @curr_change_date
				and dt > dateadd(yy, getdate(), -1)
			order by
				dt desc		
	
		    set @count = @count + 1
		    select @count, @curr_cart_num, @curr_change_date-------------
		end
	else
		begin
			set @count = 5 
			set @curr_change_date = null
		end		
  end


select dateadd(yy, @curr_change_date, 1)








 ---------------------------- тесты


	select top 1
	  old_card
	  , d
	 
	from 
	  #t
	where
	  a = 456--@curr_cart_num
		and d < 1018--@curr_change_date
	order by
	  d desc








-- рекурсия (не пошла - нужно добавить фильтр на повторения номеров внутри рекурсии)




declare @oc int = 000;

WITH r (i, a, d)
AS
(
 SELECT i, a, d
 FROM #t tt
 WHERE tt.a = @oc
 and tt.d = (select top 1 ttt.d from #t ttt where a = @oc order by d desc)
 --or not exists (select i from #t where i = @oc and d > tt.d)
 
 UNION ALL
 SELECT t.i, t.a, t.d
 FROM #t t
 JOIN r rec ON t.a = rec.i
 
 where rec.d >= t.d

)
SELECT i, a, d, lead(i) over(order by d) pr
FROM r
order by d asc


declare @oc int = 000;


 SELECT i, a, d
 FROM #t tt
 WHERE tt.a = @oc
 and tt.d = (select top 1 ttt.d from #t ttt where a = @oc order by d desc)
 --or not exists (select i from #t where i = @oc and d > tt.d)



drop  table #t


