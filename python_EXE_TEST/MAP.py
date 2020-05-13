sp = int(input("請輸入收縮壓："))
dp = int(input("請輸入舒張壓："))

MAP = (sp + (2 * dp)) / 3

if(MAP>=70 and MAP <=120):
    print("平均動脈壓為",MAP,"mmHg")
    print("為正常範圍")
else:
    print("平均動脈壓為",MAP,"mmHg")
    print("此數值須注意")