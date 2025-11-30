
### 将PPTX文件转换为AstrBot的传话筒插件使用的图片格式

![alt text](img1.jpg)

#### 用法

1. 使用PPTX排版8:3的幻灯片，用“分节”功能标识差分，示例如下![alt text](image-1.png)
2. 将PPT所有页导出png文件到当前文件夹
3. 使用这个脚本，将图片（“幻灯片1.png”、“幻灯片2.png”等）转换为AstrBot的传话筒插件使用的character差分格式
4. 将转换后的文件夹放入AstrBot的传话筒插件的“character”文件夹中
5. 在传话筒WebUI配置一个全空的预设，立绘图层在主文本框下方，铺满整个页面，设置一下文本框及字体（暂未实现自动）

#### 原理

使用PPT导出PNG图片，通过解析PPTX（实际上是个压缩包）内xml文件，将分节与导出的图片对应，重构文件夹结构后，将整张图片作为传话筒插件的立绘使用。

#### 问题

1. 显然这个是单图层的，从PPT生成多图层配置应该可行，但是较为麻烦。而python-pptx可能过老，没有section（分节）功能，使得开发很不方便。
2. 传话筒插件貌似只能是8:3的图
3. 没有设置过的PPT导出大概率会非常模糊，需要按照[如何从 PowerPoint 导出高分辨率（高 dpi）幻灯片](https://learn.microsoft.com/zh-cn/office/troubleshoot/powerpoint/change-export-slide-resolution)设置一下

---

**PPT素材来源**：Bilibili@二五母鸡卡
