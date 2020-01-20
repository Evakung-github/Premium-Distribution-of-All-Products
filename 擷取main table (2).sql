 Select a.CHDRNUM,a.crtable, a.LIFESEX,a.LIFEAGE,PCESTRM,a.OSUMINS,b.NOWPREM,b.SRCEBUS 
into [dbo].[test_1] 
from [ACDB].[dbo].[COVR] a
join  [ACDB].[dbo].[UC01] b
on
a.CHDRNUM = b.CHDRNUM
right join dbo.pro_list c
on 
a.CRTABLE = c.crtable
where 
b.STATCODE = 'IF' and b.ISSUDATE>='20170101';




select * from test_1 where crtable = '5IO3'

sp_rename 'test_1.PCESTRM','pay_year';
sp_rename 'test_1.OSUMINS','pl_sa';
sp_rename 'test_1.NOWPREM','pl_prem';
go

alter table test_1 add sadiscount float
update test_1
set test_1.sadiscount = t.discount
from test_1, discount t where test_1.pay_year = t.payment and test_1.crtable=t.crtable and test_1.pl_sa>=t.down and test_1.pl_sa<t.up
go

alter table test_1 add key_age int
update test_1
set test_1.key_age =  case when highage.highage is null then cast(lifeage/10 as int)*10
					else highage.highage
					end
					from test_1 t left join highage on t.lifeage=highage.highage and t.pay_year = highage.pay_year and t.crtable=highage.product_name
go

--update #try
--set #try.haha =  
--case when num > =65 then 65
--else cast(num/10*10 as int)
--end


alter table test_1 add saindex varchar(10)
update test_1
set saindex = cast(cast(sadiscount*100.0 as dec(18,2)) as varchar(4)) +'%'+'_'+cast(case LIFESEX when 'M' then 0 else 1 end  as varchar)
 
go 


--UPDATE list$ 
--SET len8 = com.Len8
--FROM 
--list$ 
--right JOIN com
--ON 保單號碼=com.LEN11