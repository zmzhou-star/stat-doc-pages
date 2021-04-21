# stat-doc-pages

#### 介绍
JACOB(Java-COM Bridge)统计一个文件夹下所有word文档页数

#### 软件架构说明
* [JACOB](https://sourceforge.net/projects/jacob-project/)

#### 安装教程
1. `打包`
```
mvn clean install -X -DskipTests
```
2. `运行（Windows或Linux）`
```
java -jar -server stat-doc-pages-1.0.jar
```
当前ssh窗口被锁定，可按CTRL + C打断程序运行，或直接关闭窗口，程序退出。

#### License
[The Apache-2.0 License](http://www.apache.org/licenses/LICENSE-2.0)

### 如果报错 ***jacob-1.20-x64***
请将dll文件：lib/jacob-1.20-x64.dll 放到C:\Windows\System32下，再重试
