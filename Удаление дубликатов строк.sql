
-- проверка и удаление дубликатов строк

-- тестовая таблица
create table #A(col1 int, name varchar(20), property int, d_date date )
insert into #A values (1, 'AAA', 1, '20201011')
insert into #A values (4, 'DDD', 1, '20201012')
insert into #A values (4, 'DDD', 1, '20201012')
insert into #A values (4, 'DDD', 1, '20201012')


--проверка
select distinct *, count(*) from #A
group by col1, name, property, d_date
having count(*) > 1


--удаление
alter table #A add num int

update #A
set num = %%physloc%%

delete #A
where num not in (
				select distinct max(num) over(partition by col1, name, property, d_date)
				from #A
				) 

alter table #A drop column num 


--результат

select
	 a.* 
from #A a

--drop table #A



