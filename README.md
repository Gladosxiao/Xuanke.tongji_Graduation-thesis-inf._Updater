# 选课网毕业论文相关信息更新器

为了应对xuanke.tongji.edu.cn的毕业论文管理系统不能批量添加的毕业论文信息，搞了个半自动脚本

写的很糙，但是如果后续还有助管需要大量更新信息，希望能帮大家跨出一步办公自动化的脚步

读入信息 → 打开浏览器 → 登录账户 → 进入录入界面（指定学生的编辑界面） → …

… → 输入第一个开始录入学生的学号 → 自动填表 → …

… → 人肉检查+保存 → 切换到下一位学生页面 → 显示下一位学生姓名 →回车进行下一位学生填表

_这里为了谨慎起见，没有全搞自动化，因为对要输入的excel表还有担心数据有问题的地方(老师指导的学生超上限之类的Bug)_

(╯‵□′)╯︵┻━┻啥时候能批量导入信息呀，这个老系统

## selenium有关信息

如果你想看看selenium的介绍或者wiki:

[Selenium with Python](https://selenium-python.readthedocs.io/)

 ---------- 
如果你想看看selenium的无法调用IE浏览器:

_由于选课网只支持IE+同济VPN才能更新信息，浏览器也只限制了使用IE_

检查驱动：

[selenium对应三大浏览器（谷歌、火狐、IE）驱动安装](https://www.cnblogs.com/qiezizi/p/8632058.html)

检查安全策略：

[How to run webdriver in IE browser?](https://www.seleniumeasy.com/selenium-tutorials/how-to-run-webdriver-in-ie-browser)

检查注册表设置：

[Not able to launch IE browser using Selenium2 (Webdriver) with Java](https://stackoverflow.com/questions/14952348/not-able-to-launch-ie-browser-using-selenium2-webdriver-with-java)

 ---------- 
如果你的代码无法找到xuanke.tongji.edu.cn页面上的元素:

[find_element_by_name找不到element](https://blog.csdn.net/xiaoaiwhc/article/details/8620427)

[解决套娃问题的代码 Moving between windows and frames](https://selenium-python.readthedocs.io/navigating.html)

 ---------- 
 
**愿这个世界多点自动化，少点人肉输入o(╥﹏╥)o**
