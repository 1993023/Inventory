# """
#  # for i in ed["id"]:
#     #     for j in RLAct:       # for type "L"
#     #         if (get_actual(i,j))>0:
#     #                 ed.loc[i,j]=0
#     #                 # print(ed[j])
#     # print("done")
 
#     # for i in ed["id"]:
#     #     for j in RSAct:       # for type "S"
#     #         if (get_actual(i,j))>0:
#     #                 ed.loc[i,j]=0
#                     # print(ed[j])
#     # print("done")

# """

# """
# final rollback

# RCAct=["RCActP1","RCActP2","RCActP3","RCActP4","RCActP5","RCActP6","RCActP7","RCActP8","RCActP9","RCActP10"]
# RLAct=["RLActP1","RLActP2","RLActP3","RLActP4","RLActP5","RLActP6","RLActP7","RLActP8","RLActP9","RLActP10"]
# RSAct=["RSActP1","RSActP2","RSActP3","RSActP4","RSActP5","RSActP6","RSActP7","RSActP8","RSActP9","RSActP10"]



# def get_actual(val1,val2):
#     actvalue=None
#     s1=ed
#     actvalue=s1.loc[(s1["id"]==val1),val2]
#     if len(actvalue)>0:
#         actvalue=s1.loc[(s1["id"]==val1),val2].values[0]
#     return actvalue

# # print(get_actual(1,"RLActP1"))

# def get_cancel(val1,val2):
#     cvalue=None
#     s1=ed
#     cvalue=s1.loc[(s1["id"]==val1),val2]
#     if len(cvalue)>0:
#         cvalue=s1.loc[(s1["id"]==val1),val2].values[0]
#     return cvalue

# # print(get_cancel(11,"RCActFlag"))



# def rollback(ed):
#     lst1=[]
#     for i in ed["id"]:
#         if ((get_cancel(i,"RCActFlag"))=="c" or (get_cancel(i,"RCActFlag"))=="C"):
#             lst1.append(i)
#             break
#     else:
#         lst1.append(1000)

#     r=lst1[0]
#     for k in range(r,len(ed)):
#         for j in RCAct:       # for type "C"
#             # if (get_actual(k,j))>0:
#                 # print(k,j)
#                 ed.loc[k,j]=0
#                 # print(ed[j])
#     print("ok!")


#     lst2=[]
#     for i in ed["id"]:
#         if ((get_cancel(i,"RLActFlag"))=="c" or (get_cancel(i,"RLActFlag"))=="C"):
#             lst2.append(i)
#             break
#     else:
#         lst2.append(1000)

#     r=lst2[0]
#     for k in range(r,len(ed)):
#         for j in RLAct:       # for type "L"
#             # if (get_actual(k,j))>0:
#                 ed.loc[k,j]=0
#                 # print(ed[j])
#     print("ok!")



#     lst3=[]
#     for i in ed["id"]:
#         if ((get_cancel(i,"RSActFlag"))=="c" or (get_cancel(i,"RSActFlag"))=="C"):
#             lst3.append(i)
#             break
    
#     else:
#         lst3.append(1000)
    
#     r=lst3[0]
#     for k in range(r,len(ed)):
#         for j in RSAct:       # for type "S"
#             # if (get_actual(k,j))>0:
#                 ed.loc[k,j]=0
#                 # print(ed[j])
#     print("ok!")
        
                 
#     return ed



# # rollback(ed)
# ed=rollback(ed)

# ##updating output to excel
# updatedftoexcel(ed,'Sheet1')


# """


# # for i in range(10,5):
# #     if i==7:
# #         break
# # else:
# #     print("else executed")

# """
# anothet way


# RCAct=["RCActP1","RCActP2","RCActP3","RCActP4","RCActP5","RCActP6","RCActP7","RCActP8","RCActP9","RCActP10"]
# RLAct=["RLActP1","RLActP2","RLActP3","RLActP4","RLActP5","RLActP6","RLActP7","RLActP8","RLActP9","RLActP10"]
# RSAct=["RSActP1","RSActP2","RSActP3","RSActP4","RSActP5","RSActP6","RSActP7","RSActP8","RSActP9","RSActP10"]
# RActualF=["RCActFlag","RLActFlag","RSActFlag"]


# def get_cancel(val1,val2):
#     cvalue=None
#     s1=ed
#     cvalue=s1.loc[(s1["id"]==val1),val2]
#     if len(cvalue)>0:
#         cvalue=s1.loc[(s1["id"]==val1),val2].values[0]
#     return cvalue

# # print(get_cancel(7,"RCActFlag"))




# def test(ed):
#     for i in ed["id"]:
#         for j in RActualF:
#             if get_cancel(i,j)=="c" or get_cancel(i,j)=="C":
#                 # print(i)
#                 if j =="RCActFlag":
#                     for r in range(i,len(ed)):
#                         for k in RCAct:
#                             ed.loc[r,k]=0
#                 if j =="RLActFlag":
#                     for r in range(i,len(ed)):
#                         for k in RLAct:
#                             ed.loc[r,k]=0
#                 if j =="RSActFlag":
#                     for r in range(i,len(ed)):
#                         for k in RSAct:
#                             ed.loc[r,k]=0
#     print("done")
#     return ed
                

# # test(ed)
# ed=test(ed)

# #updating output to excel
# updatedftoexcel(ed,'Sheet1')

# """


# # a=[0 if x==3 else x for x in range(0,10)]
# # print(a)



# # three types of looping using list
# # a=["a","b","c","d"]
# # for i in range(len(a)):
# #     print(i)

# # for i in range(len(a)):
# #     print(a[i])

# # for i in a:
# #     print(i)

# # import pandas as pd
# # data=[["Alex",10],["Boby",15],["Mariya",20]]
# # df=pd.DataFrame(data,columns=["Name","Age"],index=["info1","info2","info3"])
# # print(df)


# # data={"Name":["john","crysty","maria","henry"],"Age":[23,35,25,46],"country":["USA","Russia","dubai","india"]}
# # df=pd.DataFrame(data)
# # print(df)

# # for i in range(len(df)):
# #     if(df["Name"][i]=="crysty"):
# #         print("crysty age is ",df["Age"][i])

# # def test(val1,val2):
# #     tvalue=None
# #     s1=df
# #     tvalue=s1.loc[(s1["Name"]==val1),val2]
# #     if tvalue>0:
# #         tvalue=s1.loc[(s1["Name"]==val1),val2].values[0]
# #     return tvalue
# # print(test("maria","Age"))



# # for x in range(len(plants)):
# #                             discharge_days = discharge[x]
# #                             ed['R'+j+'Act'+plants[x]][i+discharge_days] = 0
# #                         print('list_---- Once ',list_)


# """
# simplified 

# RCAct=["RCActP1","RCActP2","RCActP3","RCActP4","RCActP5","RCActP6","RCActP7","RCActP8","RCActP9","RCActP10"]
# RLAct=["RLActP1","RLActP2","RLActP3","RLActP4","RLActP5","RLActP6","RLActP7","RLActP8","RLActP9","RLActP10"]
# RSAct=["RSActP1","RSActP2","RSActP3","RSActP4","RSActP5","RSActP6","RSActP7","RSActP8","RSActP9","RSActP10"]
# RActualF=["RCActFlag","RLActFlag","RSActFlag"]

# def rollback(ed):
#     for i in range(len(ed)):
#         if ed['RCActFlag'][i] == 'c' or ed['RCActFlag'][i] == 'c':
#             for k in RCAct:
#                 ed[k][i:len(ed)] = np.NAN
#         if ed['RLActFlag'][i] == 'c' or ed['RLActFlag'][i] == 'c':
#             for k in RLAct:
#                 ed[k][i:len(ed)] = np.NAN
#         if ed['RSActFlag'][i] == 'c' or ed['RSActFlag'][i] == 'c':
#             for k in RSAct:
#                 ed[k][i:len(ed)] = np.NAN

        
#     print("done")
#     return ed
                

# # rollback(ed)
# ed=rollback(ed)

# #updating output to excel
# updatedftoexcel(ed,'Sheet1')


# """          




# """
# modified for "C" 

#     #                 if str(ed['R'+j+'ActFlag'][i])=="c" or str(ed['R'+j+'ActFlag'][i])=="C" :
#     #                     if j == 'C':
#     #                         list_ag = ed['R'+j+'Agg']
#     #                         print('Value',ed['R'+j+'ActFlag'][i],'Typej=',j,' Plantk=',k,' Rowi=',i)                 

#     #     #                     ed['R'+j+'ActFlag'][i] = def_inventory_g 
#     #                         t = checkNegative1(list_ag,i,ed['R'+j+'ActFlag'][i])
#     #                         print("######################",t)
#     #                         list_ = contriPlant(ed,j,i,t+1)
#     #                         # print("*********************************************************************",list_)
#     #                         # list_.clear()
#     #                         print("*********************************************************************",list_)
#     #                         for x in range(len(plants)):
#     #                             discharge_days = discharge[x]
#     #                             ed['R'+j+'Act'+plants[x]][i+discharge_days] = 0 
#     #                         print('list_---- Once ',list_)


#     #     #                     ed['R'+j+'Act'+k][i] = ed['R'+j+k][i:t].fillna(0).sum()

#     #                         ed['R'+j+'ActFlag'][t+1] =  def_inventory_g
#     #                         m = t + 1
#     #                         while m <= len(ed):
#     #                             print('m',m)
#     #                             #[m+1:]

#     #                             a = checkNegative1(list_ag,m,def_inventory_g)

#     #     #                         a ,list_ = checkNegative1(ed,j,list_ag,m,def_inventory_g)

#     #                             print('a-----',a)
#     #                             if a == None :
#     #                                 m = len(ed) + 1
#     #                             else : 
#     # #                                 list_ = contriPlant(ed,j,m,a+1)
#     # #                                 for x in range(len(plants)):
#     # #                                     ed['R'+j+'Act'+plants[x]][m] = list_[x] 
#     #                                 ed['R'+j+'ActFlag'][a+1] = def_inventory_g
#     #                                 m = a + 1

#     #                         temp = 'Found'
#     #                         #print('break3-----------')
#     # #                         print('**********','R'+j+'Flag',len(ed['R'+j+'Flag'][i+1:len(ed)]),'R'+j+'ActFlag',len(ed['R'+j+'ActFlag'][i+1:len(ed)]))
#     #                         ed['R'+j+'Flag'][i+1:len(ed)] = ed['R'+j+'ActFlag'][i+1:len(ed)]
#     #                         temp1 = ed['R'+j+'ActFlag'][0:i+1] 
#     #                         ed['R'+j+'ActFlag'] = np.NAN
#     #                         ed['R'+j+'ActFlag'] = temp1
#     #                         break
                        



# from datetime import date
# today = date.today()
# d1 = today.strftime("%d-%m-%Y")
# print(d1)



# import datetime

# Current_Date = datetime.datetime.today()
# print ('Current Date: ' + str(Current_Date))

# Previous_Date = datetime.datetime.today() - datetime.timedelta(days=1)
# print ('Previous Date: ' + str(Previous_Date))

# NextDay_Date = datetime.datetime.today() + datetime.timedelta(days=1)
# print ('Next Date: ' + str(NextDay_Date))

l1=[1,3,4,5,67,23,26,25,10]

# l2=[x for x in l1 if x%2==0 if x%5==0]
# print(l2)

# l2=[x*y for x in l1 ]
# print(l2)





