#pipConfig
import os


def pipSource(url):
	text = f"[global]\nindex-url = {url}"
	appdata = os.getenv("APPDATA")
	path = os.path.join(appdata,"pip")
	if not os.path.exists(path):
		os.makedirs(path)
	file = os.path.join(path,"pip.ini")
	if not os.path.exists(file):
		print("未发现pip源 配置文件")
		with open(file,"wt",encoding = "utf-8") as fileop:
			fileop.write(text)
		print(f"已设置pip源为 {url}")
	else:
		print("pip源 已配置")

def main():
	url = "https://pypi.tuna.tsinghua.edu.cn/simple/"
	print("*"*36)
	pipSource(url)
	print("*"*36)
	os.system("pip install pandas")
	os.system("pip install xlrd")
	os.system("pip install openpyxl")
	
	print("*"*36)
	input("pandas安装完成，可以开始运行成绩分析程序。")
if __name__ == '__main__':
	main()