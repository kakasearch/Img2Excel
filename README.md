# Img2Excel
 将照片填充进Excel中的VB版本实现

原理为读取照片像素，将每个像素的的颜色填充进Excel单元格中，并将单元格设置为正方形。

## 使用

1. 在Excel中导入模块

2. 运行宏 ```照片像素填充单元格```

3. 选择转化的照片

4. 静待片刻，完成填充

   

## demo

1. 打开文件./test.xlsm
2. 按下快捷键 ```ALt+F8```，选择运行宏 ```照片像素填充单元格```
3. 选择./test_100x100.jpg
4. 等待片刻，完成填充

## 其他

1. png照片需要先将透明部分填充成白色，否则程序会报错
2. 照片尺寸建议小于200*200。尺寸越大，运行越慢
