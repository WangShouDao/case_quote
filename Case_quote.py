import numpy as np
import itertools, time, xlrd

A = []	# 所有的评标价
B = []	# 公司的评标价
C = []	# 备选的评标价

def read_excel():
	path= 'F:\python\示范案列报价.xls'
	workbook = xlrd.open_workbook(path)
	sheets = workbook.sheet_names();
	worksheet = workbook.sheet_by_name(sheets[0])
	
	#获取数据
	j = 12	# 第十三列
	for i in range(1, 32):	# 二到三十二行
		A.append(worksheet.cell_value(i,j))	
	
	for i in range(13, 17):
		B.append(worksheet.cell_value(i,j))	
	
	for i in range(32, 48):
		C.append(worksheet.cell_value(i,j))

	return (A,B,C)
	
# 招标控制价
S = 364157469.00 

# 第一轮废标（评标价小于招标控制价相应价格*15%）
r1 = 0.85

# 第二轮废标1（评标下浮10%）
r2 = 0.9

# 第二轮废标2（评标价平均值的95%）
r3 = 0.95


# 清楚第一轮中不符合条件的投标
def clean_one(Y):		
	M1 = S * r1	# 第一轮废标中的最低价
	Y = [ x for x in Y if x >= M1]		
	return Y

# 计算第二轮评标条件的平均值、评标下浮	
def calc_Avg(Y):
	Avg = np.average(Y) * r3
	M2 = S * r2	
	M = min(Avg, M2)
	return (Avg, M2, M)

# 计算集合中满足第二轮评标条件的最小值
def calc_min(X, Y):
	y=0
	Avg, M2, M = calc_Avg(X)
	if len(list(filter(lambda x : x >= M, Y))) != 0:
		y = np.min(list(filter(lambda x : x >= M, Y)))
	return y
	
# 对集合中的数据排序后进行组合	
def combination_list(Y):
	list1 = []
	Y = sorted(Y)
	for i in range(1,len(Y)+1):
		iter = itertools.combinations(Y, i)
		list1.extend(iter)
	return list1

def description(A0, B0, Avg, M2):
	print("所有评标价的Avg值为: %f, 评标下浮的值M2为: %f" % (Avg, M2))	
	print("所有评标价中满足条件的最小值A0为: %f , 与中标条件的差值为： %f" % (A0, A0 - Avg))
	print("公司评标价中满足条件的最小值B0为: %f , 与中标条件的差值为： %f" % (B0, B0 - Avg))
	
def add_price(A, B, C):
	A0 = calc_min(A, A)
	B0 = calc_min(A, B)
	Avg, M2, M = calc_Avg(A)
	print('------------------------------')
	description(A0, B0, Avg, M2)
	# 如果B0、A0满足B0 <= A0，则成功中标，程序结束
	if B0 and A0 and B0 <= A0:		
		print("不用添加新的投标即可成功中标，中标价格为：%f" % B0)
		return
	# 对集合C中的所有元素进行组合，并依次添加到集合A、B中进行计算
	else:
		print("需要添加备选方案：")
		print('------------------------------')
		list1 = combination_list(C)
		k = 1
		list2 = []
		# 对集合C所有的组合方案进行遍历
		for i in range(len(list1)):
			A.extend(list(list1[i]))
			B.extend(list(list1[i]))			
			A0 = calc_min(A, A)
			B0 = calc_min(A, B)
			Avg, M2, M = calc_Avg(A)		
			if B0 and A0 and B0 <= A0:
				if list(list1[i]) not in list2:
					list2.append(list(list1[i]))
					print("推荐的添加方案",k,"为:",list(list1[i]))
					description(A0, B0, Avg, M2)				
					print("添加推荐方案后可以成功中标，中标价格为: %f" % B0)
					print('------------------------------')
					A = A[:(len(A)-len(list1[i]))]
					B = B[:(len(B)-len(list1[i]))]
					k += 1
				else:
					A = A[:(len(A)-len(list1[i]))]
					B = B[:(len(B)-len(list1[i]))]
			else :
				A = A[:(len(A)-len(list1[i]))]
				B = B[:(len(B)-len(list1[i]))]
				
			if k > 10 :
				return
		if k == 1 :
			print("不存在满足调价的投标方案可以添加，投标失败")
			return
		
if __name__ == "__main__":
	start = time.time()
	read_excel()
	A,B,C = clean_one(A),clean_one(B),clean_one(C)
	add_price(A, B, C)
	end = time.time()
	print("共用时 %f s" % ((end - start)))