# -*- coding:cp936 -*-
'''------------------------------------------------------
【程序改错】
---------------------------------------------------------

题目：一家商场在降价促销。如果购买金额为50~100元(包含50元和100元) ，就会有10%的
      折扣；如果购买金额大于100元，就会有20%的折扣。编写程序，询问购买价格，再
      显示出折扣(10%或20%)和最终价格。
------------------------------------------------------------
注意：不可以增加或删除程序行，也不可以更改程序的结构。
------------------------------------------------------'''
customer_price=float(input("please input pay money:"))
#**********FOUND**********
if 100=>customer_price>= 50:
    print("disconunt 10% , after discount you shoud pay {}". format(customer_price*(1-0.1)))
#**********FOUND**********
else customer_price >100:
#**********FOUND**********
    print("disconunt 20% , after discount you shoud pay {}". format(customer_price*0.2))
else:
    print("disconunt 0% , after discount you shoud pay  {}". format(customer_price))