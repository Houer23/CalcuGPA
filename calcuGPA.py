import pandas as pd
import os
import time

#datdFrame， 文件内容
df = []

def CheckFile(filename,tail = ".txt", pre = "", suf = "", contain = ""):
	'''
	使用规范：tail参数应带"."号
	'''
	#检验文件扩展名
	if not filename.endswith(tail):
		return False
	#检验文件名开头标识
	if pre and not filename.statswith(pre):
		return False
	#检验文件名结束标识（不包含扩展名）
	testname = filename.replace(tail,"")
	if suf and not testname.endswith(suf):
		return False
	#检验文件名关键字：
	if contain and not contain in filename:
		return False
	#文件名符合所需检测的相应项目要求
	return True

def ListFiles(opath):
	allfile=[]
	pathfiles=os.listdir(opath)
	print("*"*36)
	print("序号\t文件名")
	for file in pathfiles:
		if CheckFile(file,tail = ".xlsx"):
			ifile=os.path.join(opath,file)
			allfile.append(ifile)
			print("{:0>2}){}".format(len(allfile),file))
	print("*"*36)
	return allfile

def GetFile(path):
	files=ListFiles(path)
	while True:
		s=input("请选择文件，输入其所对应的序号：")
		try:
			num=int(s)-1
			filename=files[num]
		except:
			print("\n<{}>序号输入错误！！！\n".format(s))
		else:
			break
	return filename

def prInfo(text,sleepTime = 0):
	print(text)
	time.sleep(sleepTime)


#以下开始为核心处理内容
def SingleResult(num):
	dfi = df.loc[df["学号"] == num, ["课程代码","学分","绩点"]]
	subjects = {}
	for index in dfi.index:
		subjects[dfi.loc[index,"课程代码"]] = (dfi.loc[index,"学分"], dfi.loc[index,"绩点"])
	all_score = sum([s[0] * s[1] for s in subjects.values()])
	all_credit = sum([s[0] for s in subjects.values()])
	result = round(all_score/all_credit,5)
	return result

def MainWork(file):
	global df
	df = pd.read_excel(file)
	nums = set(df.loc[:,"学号"])
	prInfo(f"已读取学号列表，总共{len(nums)}位同学",3)
	#开始处理数据
	students = {}
	for num in nums:
		print(f"学号：{num}，",end = "")
		singleCredit = SingleResult(num)
		print(f"GPA：{singleCredit}")
		students[num] = singleCredit
	for index in df.index:
		df.loc[index, "A"] = students[df.loc[index,"学号"]]
	return 1
#核心代码如上。


def Result(filepath):
	path,file = os.path.split(filepath)
	filename,tail = os.path.splitext(file)
	newname = f"{filename}-result{tail}"
	result_file = os.path.join(path,newname)
	return result_file

def main():
	#选择文件夹
	while True:
		print("*"*36)
		path = input("请输入表格所在文件夹的绝对路径：")
		if os.path.exists(path):
			break
		else:
			print("文件夹不存在，请检查输入是否正确！！！")
	#假定输入文件夹正确，选择文件
	file = GetFile(path)
	print("*"*36)
	print("开始处理文件\n")
	s = MainWork(file)
	if s:
		result_file = Result(file)
		print("成绩处理完成")
		df.to_excel(result_file,index = False)
		input(f"已保存至：\n\t{result_file}")
	else:
		input("可能出错了。。。")

if __name__ == '__main__':
	main()

