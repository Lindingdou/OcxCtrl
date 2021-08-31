# :star:学习笔记

##  :one: 引言

ActiveX是Microsoft对于一系列策略性面向对象程序技术和工具的称呼，其中主要的技术是组件对象模型（COM）。 ActiveX控件是一种实现了一系列特定接口而使其在使用和外观上更象一个控件的COM组件。ActiveX控件这种技术涉及到了几乎所有的COM和OLE的技术精华，如可链接对象、统一数据传输、OLE文档、属性页、永久存储以及OLE自动化等。

- ActiveX控件作为基本的界面单元，必须拥有自己的属性和方法以适合不同特点的程序和向包容器程序提供功能服务，其属性和方法均由自动化服务的 IDispatch接口来支持。
- 除了属性和方法外，ActiveX控件还具有区别于自动化服务的一种特性--事件。事件指的是从控件发送给其包容程序的一 种通知。
- 与窗口控件通过发送消息通知其拥有者类似，**ActiveX 控件**是通过触发事件来通知其包容器的。
- 事件的触发通常是通过控件包容器提供的**IDispatch接口**来调用自动化对象的方法来实现的。
- ActiveX控件的方法、属性和事件均有自定义（**custom**）和库存（**stock**）两种不同的类型：
  - 自定义（**custom**）的方法和属性也就是是普通的自动化方法和属性，自定义事件则是自己选取名字和Dispatch ID的事件。
  - 库存（**stock**）的方法、属性和事件则是使用了ActiveX控件规定了名字和Dispatch ID的"标准"方法、属性和事件。

## :two: MFC ActiveX项目向导配置

MFC ActiveX控件向导参数配置说明：

![](https://github.com/Lindingdou/OcxCtrl/blob/master/image/classview.jpg?raw=true)

- CoxControlApp 是控件的主程序模 块 ， 定义了控件的注 册（ DllRegisterServer）、删除（DllUnregisterServer）等功能，一般不用动，如有需要可以在其中的 InitInstance 和ExitInstance 中定义我们自己的初始化和终止操作代码，一般也就是一些资源的初始化和销毁工作。
- CoxControlCtrl是控件类，要做的控件功能基本上就是要在这个类中实现。需要提一下的是在这个类中重写了父类的 OnDraw 函数。实际控件开发中可以根据功能需要修改重写这个函数来绘制控件界面。
- CoxControlPage 是属性页类，这个类实现了一个在开发时设定控件属性的对话框。
- CoxControlLib 是为客户程序提供本控件的属性、方法以及可能响应的事件的接口的。

## :three: 属性、方法和事件

属性：属性是AcitveX控件想所有容器公开的数据成员。

- 控件实现属性的目的是可以自定义它们来满足特殊应用程序或Web网页的需求。

方法：方法就是控件开放给用户使用的一些功能函数，类似于C++的类函数。开发人员可以提供自定义功能。

- 控件实现方法是为了调用他们执行有用的工作。

事件：ActiveX通过事件通知容器控件上发生了某些事情。将控件开发人员的某一特定操作识别为事件。

- 容器是指拥有ActiveX控件的窗口。
- 窗口控件通过发送消息给它们的所有者发出通知；ActiveX控件通过激发事件给他们的容器发出通知。
- 事件的激发是通过控件容器提供的接口（通常是IDispatch接口）调用Automation方法来实现的。
- ActiveX控件规范中的一部分内容就用来解释了控件怎么获得指向容器的IDispatch接口的指针。

## :four: 自定义型和备用型

方法、属性和事件分为两种类型：自定义型和备用型

- 自定义型：可以选择它们的名字并分配调度ID
- 备用型：使用ActiveX控件规范中指定的名称和调度ID

## :five:  环境属性

环境属性是一种Automation属性，它是通过IDispatch公开的，不同之处在于是容器而不是控件实现了接口。

环境属性具有名称和调度ID

## :six:  控件状态

控件两种状态：活动和不活动

- 活动的控件时处于活动状态，在容器中运行着的控件
- 不活动的控件不具有自己的窗口，它会小号更少的系统资源并且不响应用户输入。

## :seven:  ActiveX控件体系结构

:key:“MFC Windows程序设计”书中图21-2图中所描绘的控件为MFC创建的ActiveX控件。

- 控件对象保存在一个Win32 DLL中，它通常被称为OCX，其中，OC代表OLE Control。
- 一个OCX可以保存一个以上的控件
- ActiveX控件是一种在窗口中绘制可见表现形式的控件。
- 如果需要激发事件，需要实现IConnectionPointContainer接口
  - IConnectionPointContainer接口间接地促使容器给控件提供IDispatch接口指针
  - 激发事件，控件会利用容器的IDispatch指针调用Automation方法
- 如果需要公开方法和属性，则需要实现IDispatch接口

## :eight: ActiveX控件容器

为了拥有ActiveX控件，容器必须实现COM接口

## :nine:  MFC对ActiveX控件的支持

MFC简化了ActiveX控件和控件容器的编写，提供了所要求的COM接口的实现，封装了连接ActiveX控件和控件容器的协议。

- 对于那些不能以一般方式实现的COM方法，MFC提供了虚函数，您可以覆盖它们实现控件特定的功能。
- MFC编写ActiveX控件就是从MFC基类中派生出类并在某些地方覆盖其虚函数，添加程序逻辑结构使控件与别的控件不同。

:angel:对于部分MFC类的介绍：

- COleControl

  - 所有MFC ActiveX控件的基类，它实现了大多数COM接口。
  - 它包含许多Windows消息的处理程序，并提供了内置的备用方法、属性和事件的实现。
  - 共有60多个虚函数，代表性的有：

    - onDraw：用来绘制控件，在添加控件特有的绘图逻辑程序时覆盖它

    - DoPropExchange：用来保存和加载控件的持久性属性。为了支持持久性控件和属性而覆盖它。
  - 非虚函数：
    - InvalidateControl:重绘控件
- COleControlModule
  - 主要贡献是InitInstance函数，它调用AfxOleInitModule使COM在MFC DLL中得到支持
  - 如要覆盖COleControlModule派生类的InitInstance函数，就必须在执行自己的任何程序之前调用基类的InitInstance函数
  - 在ControlWizard创建ActiveX控件项目时，它已经添加了调用。
- COlePropertyPage
  - 通过实现属性表，ActiveX控件会给开发者公开它们的属性，该属性表由控件容器显示
  - 属性页是COM对象，又称OLE属性页，是实现了COM的IPropertyPage和IPropertyPage2接口的对象
- CConnectionPoint和COleConnPtContainer
  - 为了进行事件激发，ActiveX控件使用了COM的可连接对象协议，以从它们的容器中接受接口指针。
  - 可连接对象是指实现了一个以上“连接点”的对象。
  - 实际上，连接点是一个实现了IConnectionPoint接口的COM对象。
  - 为了将连接点公开，可连接对象是实现了COM的IConnectionPointContainer接口。
- COleControlContainer和COleControlSite
  - IBoundObjectSite
  - INotifyDBEvents
  - IRowsetNotify

## :keycap_ten: 创建ActiveX控件

ControlWizard步骤及选项详解：

- 运行时许可证
- 创建的控件基于
- 可见时激活
- 运行时不可见
- 有”关于“对话框
- 优化的绘图代码
- 无窗口激活
- 未剪辑的设备上下文
- 无闪烁激活
- 在”插入对象“对话框中可用
- 不活动时有鼠标指针通知
- 作为简单框架控件
- 异步加载属性

### :arrow_right: 1. 实现OnDraw



### :arrow_right: 2. 使用环境属性



### :arrow_right: 3. 添加方法



### :arrow_right: 4. 添加属性



### :arrow_right: 5. 使属性成为持久属性



### :arrow_right: 6. 自定义控件的属性表



### :arrow_right: 7. 给控件属性表添加页



### :arrow_right: 8. 添加事件



###  :arrow_right: 9. 事件映射表









