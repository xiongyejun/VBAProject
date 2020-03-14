# 利用VBAProject来共用VBA代码

VBA代码是随文件一起保存的，个人一直以来，使用都是在某一个文件里来编辑代码。

随着VBA使用的增多，必然会积累一些常用的代码，甚至在网上也会找到一些功能强大的类。每次使用也都是复制到某个文件里使用，这些代码在自己电脑里就存在许多个副本。使用过程中难免会发现一些问题，对代码进行一些小的修改，一些类也可能增加一些实用的方法、函数。但是修改了其中一个副本，电脑里其他使用了的文件没法一同更改。久而久之，各个副本中就会存在不同时间里修改过的代码，想把他们放一起会发现变得非常困难。

自己有时候就希望如果VBA能像C语言那样有**#include &lt;xx&gt;** 这种方式就好了，所有程序都引用的是同一个文件里的代码，只需修改一处即可。

后来发现其实利用**工具-引用** VBAProject也能达到类似效果，**缺点是想把做好的程序文件发给别人时不大方便**。使用方法比较简单：

 1. 创建1个加载宏，修改VBAProject属性里的工程名称，以保证是唯一的，如vbapTest。使用加载宏的目的只是为了不在前台显示出文件。
 2. 将一些常用的代码保存在这个加载宏中，需要对外公开的用Public修饰，也可以省略修饰。
 3. 在需要使用这些代码的文件里，添加**工具-引用**，浏览找到加载宏，注意选择文件类型（默认是**olb,tlb,dll**，这样看不到xla和xlam，选所有文件或者Mircosoft Excel Files）。
 4. 这样只要在需要使用代码的地方，加上vbapTest.就可以调用函数或者方法了。

## 类的使用

用上述方法在碰到类的时候，会发现根本无法定义、创建类，类的使用方法需要再做1点工作，有3种方法：

 **1. 用自定义数据类型封装一下**
 在vbapTest里增加1个自定义的类型和一个函数：

        Type Test
            c As CTest
        End Type
        
        Function NewCTest() As CTest
            Set NewCTest = New CTest
        End Function
        
在使用代码的文件声明变量a为vbapTest.Test，并创建类

        set a.c = vbapTest.NewCTest()
        
然后就可以像使用同1个文件的类一样使用了。

 **2. 设置类属性Instancing**
 
 类模块有1个叫做Instancing的属性，默认是1-Private，还有1个是2-PublicNotCreatable（字面理解：公开但是不能被创建），设置为2后，在其他文件中可以声明，但不能创建，使用方法：
        
        dim a as vbapTest.CTest
        set a = vbapTest.NewCTest()
        
 **3. 强制设置类属性Instancing为5-MultiUse：**
        
    ThisWorkbook.VBProject.VBComponents("CTest").Properties("Instancing") = 5

这种方法设置过后，其他文件就完全像是使用同1个文件的类了。

    Dim c As vbapSpace.CTest
    Set c = New vbapSpace.CTest
    
这是在[网上看到的][1]，至于为什么故意不公开这个属性5，是不是会出现什么问题，目前不知道。


  [1]: http://blog.sina.com.cn/s/blog_bbfa8f220101d214.html