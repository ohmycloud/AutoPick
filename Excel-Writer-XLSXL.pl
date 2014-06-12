----------------------------------------------------------------------------
 翻译: 小Perl
Email:sxw2k@sina.com
----------------------------------------------------------------------------



NAME     名称
    Excel::Writer::XLSX - 以Excel2007+XLSX格式创建一个新文件.

VERSION  版本
    该文档是在2013年11月发布的Excel::Writer::XLSX 0.75版本。
SYNOPSIS 概要
    在perl.xlsx的第一个工作表中写入字符串、格式化的字符串、数字和公式：
        use Excel::Writer::XLSX;

        # 新建excel工作簿
        my $workbook = Excel::Writer::XLSX->new( 'perl.xlsx' );

        # 增加一个工作表
        $worksheet = $workbook->add_worksheet();

        #  添加并定义一个格式        
        $format = $workbook->add_format();#增加一种格式
        $format->set_bold();              #设置粗体
        $format->set_color( 'red' );      #设置颜色
        $format->set_align( 'center' );   #设置对齐方式（此处为居中）

		#写入一个格式化和非格式化的字符串，使用行列表示法。
        $col = $row = 0;                  #设置行和列的位置
        $worksheet->write( $row, $col, 'Hi Excel!', $format );
        $worksheet->write( 1, $col, 'Hi Excel!' );

        #使用A1表示法写入一个数字和公式
        $worksheet->write( 'A3', 1.2345 );         #在第三行第一列写入一个数字
        $worksheet->write( 'A4', '=SIN(PI()/4)' ); #在第四行第一列写入一个公式

DESCRIPTION 说明
    The "Excel::Writer::XLSX" 模块可以被用做建立Excel2007+ XLSX格式的文件。

    XLSX格式是Excel 2007 和以后版本使用的官方开放XML（OOXML）格式
    可在工作簿中添加多张工作表，格式可以被应用到单元格中。可以把文字，数字，和公式写入单元格。

   此模块目前还不能被用于向一个已经存在的EXCEl XLSX文件中写入数据。

Excel::Writer::XLSX and Spreadsheet::WriteExcel
   
	Excel::Writer::XLSX使用和Spreadsheet::WriteExcel模块相同的接口来生成二进制XLS格式的Excel文件

    Excel::Writer::XLSX 支持所有Spreadsheet::WriteExcel中的特性，并且某些情况下功能更强。请查看 "Compatibility with Spreadsheet::WriteExcel".获取更多细节。

    
    XLSL格式相比XLS格式主要的优势是它允许在一个工作表中容纳更大数量的行和列。
QUICK START 快速入门
     Excel::Writer::XLSX试图尽可能的提供Excel的功能接口。因此，有很多与接口有关的文档，第一眼很难看出哪些重要，哪些不重要。所以对于你们这些更喜欢先组装宜家设备，再读说明书的人，此处有三种简单的方式：
	0、新建一个Excel对象
    1、使用"new()"方法创建一个新的Excel 工作簿
    2、使用"add_worksheet()"方法向新工作簿增加一个工作表
	3、使用"write()"方法向工作表中写入数据
    就象这样：

        use Excel::Writer::XLSX;                                   # Step 0

        my $workbook = Excel::Writer::XLSX->new( 'perl.xlsx' );    # Step 1
        $worksheet = $workbook->add_worksheet();                   # Step 2
        $worksheet->write( 'A1', 'Hi Excel!' );                    # Step 3
    
	这会创建一个叫做perl.xlsx的Excel文件，里面只有一张工作表并在相关单元格里面有'Hi Excel'文本。
    
   
工作簿方法
       Excel::Writer::XLSX模块为新建的Excel工作簿提供了面向对象的接口。下面的方法可以通过一个新建的工作簿对象访问.

        new()                       #新建
        add_worksheet()             #添加工作表
        add_format()                #添加格式
        add_chart()                 #添加图表
        close()                     #关闭工作簿
        set_properties()			#设置属性
        define_name()				#定义名称
        set_tempdir()				#设置临时文件夹
        set_custom_color()			#设置自定义颜色
        sheets()					#工作表
        set_1904()                  #设置纪元开始年
        set_optimization()          #设置优化

  new()   
    使用"new()"构造方法创建一个新的Excel工作簿，该方法接受一个文件名或文件句柄作为参数。下面的例子根据一个文件名来创建一个新的Excel文件：
	
        my $workbook  = Excel::Writer::XLSX->new( 'filename.xlsx' );
        my $worksheet = $workbook->add_worksheet();
        $worksheet->write( 0, 0, 'Hi Excel!' );

    下面是使用文件名作为new()方法参数的其他例子：
	
        my $workbook1 = Excel::Writer::XLSX->new( $filename );
        my $workbook2 = Excel::Writer::XLSX->new( '/tmp/filename.xlsx' );
        my $workbook3 = Excel::Writer::XLSX->new( "c:\\tmp\\filename.xlsx" );#Windows
        my $workbook4 = Excel::Writer::XLSX->new( 'c:\tmp\filename.xlsx' );

    最后两个例子说明了怎样通过转义目录分隔符"\"或使用单引号保证值不被内插来在DOS上或Windows上建立Excel文件。
	
    我们推荐文件名使用".xlsx"而不是".xls"后缀，因为后者在使用XLSX格式的文件时会发生警告。
  
  
	"new()"构造函数方法返回一个Excel::Writer::XLSX对象，你可以使用这个对象来添加工作表并存储数据。	应该注意的是，尽管没有特别要求使用"my"，但是它定义了新工作簿变量的作用域，并且，在大多数情况下，它保证了工作簿不用显式地调用"close()方法"就能被正确地关闭。
	
    如果文件不能被创建，由于文件权限或其他一些原因，"new"会返回"undef"。因此，在继续之前检查"new"的返回值是个好习惯。通常，如果存在文件创建错误，Perl变量$!就会被设置：

        my $workbook = Excel::Writer::XLSX->new( 'protected.xlsx' );
        die "Problems creating new Excel file: $!" unless defined $workbook;
		
	你也可以传递一个合法的文件句柄给"new()"构造函数。例如在一个CGI程序中你可以这样做：

        binmode( STDOUT );
        my $workbook = Excel::Writer::XLSX->new( \*STDOUT );

    
    对于CGI程序，你也可以使用特别的Perl文件名 '-'，它会把输出重定向到标准输出：
        my $workbook = Excel::Writer::XLSX->new( '-' );

   可以查看例子中的cgi.pl

    然而，这种特殊的情况在"mod_perl"程序中不起作用，你必须做一些下面的事情：
        # mod_perl 1
        ...
        tie *XLS, 'Apache';
        binmode( XLSX );
        my $workbook = Excel::Writer::XLSX->new( \*XLSX );
        ...

        # mod_perl 2
        ...
        tie *XLSX => $r;    # Tie to the Apache::RequestRec object
        binmode( *XLSX );
        my $workbook = Excel::Writer::XLSX->new( \*XLSX );
        ...

    请查看mod_perl1.pl" 和 "mod_perl2.pl"
  
  
	如果你想通过socket 去 stream一个Excel文件或者你想把一个Excel文件存进一个标量，
	那么文件句柄会很有用。
	例如，下面是把Excel文件写入标量的一种方法：
        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        open my $fh, '>', \my $str or die "Failed to open filehandle: $!";

        my $workbook  = Excel::Writer::XLSX->new( $fh );
        my $worksheet = $workbook->add_worksheet();

        $worksheet->write( 0, 0, 'Hi Excel!' );

        $workbook->close();

        #  Excel文件现在在$str中。记得在打印$str之前binmode()输出文件句柄
        binmode STDOUT;
        print $str;

    请查看"write_to_scalar.pl" 和 "filehandle.pl" 

    注意binmode()的要求。一个Excel文件由二进制数据组成。因此，如果你使用一个文件句柄你应该保证在传递给"new()"方法之前将它binmode().不管你使用Windows或其他何种平台，你都应该这样做。
    
    如果你使用文件名而非文件句柄，你不用担心"binmode()".当Excel::Writer::XLSX将文件名转换为文件句柄时，它会在内部执行binmode().
---------------------------------------------------------------------------------------------------------	
  add_worksheet( $sheetname )   #参数为工作表名
   
	至少一个工作表应该被添加到工作簿中。工作表被用于将数据写入单元格：
        $worksheet1 = $workbook->add_worksheet();               # Sheet1
        $worksheet2 = $workbook->add_worksheet( 'Foglio2' );    # Foglio2
        $worksheet3 = $workbook->add_worksheet( 'Data' );       # Data
        $worksheet4 = $workbook->add_worksheet();               # Sheet4


    如果没有指定$sheetname，默认会使用Sheet1,Sheet2....
    
	
    工作表名必须是合法的Excel工作表名。它不能包含以下字符,"[ ] : * ? / \" ,并且长度必须小于32个字符。此外，你不能在一个以上的工作表中使用同一个文件名，或大小写敏感的文件名。
	
----------------------------------------------------------------------------------------------
  add_format( %properties )   #添加格式
    
	    "add_format()"方法可以被用作创建新的格式对象，它可以被用于将格式应用到单元格中。你可以在创建时通过含有属性值的哈希定义属性或之后通过方法调用定义属性：
		
	$format1 = $workbook->add_format( %props );    # 在创建时定义属性
    $format2 = $workbook->add_format();            # 创建后再定义属性

    请查看“单元格格式化”章节获取详细信息
---------------------------------------------------------------------------------------------------------
  add_chart( %properties )  #添加图表
    	该方法用于新建一个图表作为一个独立的工作表（默认），或作为一个可嵌入的对象，通过"insert_chart()"工作表方法插入到工作表中。
	
    my $chart = $workbook->add_chart( type => 'column' );

    属性可以设置为:

        type     (required)#必须的选项
        subtype  (optional)#图表子类型(可选)
        name     (optional)
        embedded (optional)

    *   "type" 类型

        这是必须的参数。它定义了将被创建的图表的类型。
         my $chart = $workbook->add_chart( type => 'line' );

        可用的type类型如下:

            area		#区域图
            bar			#条形图
            column		#柱形图
            line		#线图
            pie         #饼图
            scatter     #散点图
            stock		#存量图

    *   "subtype"       #图表类型（子类型）

        
        用于在需要时定义一个图表的子类型
            my $chart = $workbook->add_chart( type => 'bar', subtype => 'stacked' );

        
        目前只有条形图和柱形图支持子类型(stacked and percent_stacked)
		
    *   "name"

        为图表设置名字。名字属性是可选的，如果它不被支持，它会默认为"Chart1,Chart2....Chartn".
		图表名必须是合法的表名，与"add_worksheet()"规则一样。"name"属性可以在嵌套的图表中省略。
            my $chart = $workbook->add_chart( type => 'line', name => 'Results Chart' );

    *   "embedded"

        指定图表对象会通过"insert_chart()"工作表方法插入到工作表中。如果没有设置这个标志就尝试插入图表，会出现错误。
            my $chart = $workbook->add_chart( type => 'line', embedded => 1 );

            # Configure the chart.
            ...

            # 将图表插入到工作表中
            $worksheet->insert_chart( 'E2', $chart );

	查看Excel::Writer::XLSX::Chart获取更详细的关于在创建后如何配置图表对象的信息。也可查看chart_*.pl程序。


  close()
        一般地，当你的程序结束或工作簿对象超出作用域时，你的Excel文件会被自动关闭。然而，你可以使用close()方法显式地关闭Excel文件。
        $workbook->close();#显示地关闭Excel

       如果Excel文件必须在对其执行一些外部动作诸如复制、读取大小或者把它作为电子邮件的附件之前关闭，需要显式地用close()声明
        此外，"close()"被用于阻止Perl的垃圾回收器以错误的顺序处理工作簿、工作表和格式对象。这种情况出现在下面：
        如果"my()"没有被用于声明使用"new()"创建的工作簿变量作用域
        如果在子例程中调用"new()", "add_worksheet()" 或者 "add_format()"方法。
		
  
    原因是Excel::Writer::XLSX依赖Perl的"DESTROY"机制依特定顺序触发destructor析构方法。当工作簿、工作表和格式变量不是词法作用域或它们拥有不同的词法作用域时，前面这种情况不会发生。
    一般地，如果你创建一个0字节的文件或者你不能建立一个文件，你需要调用"close()"方法。
    
	  "close()"的返回值与perl关闭使用"new()"方法建立的文件的返回值一样。这允许你以常规方式处理错误
        $workbook->close() or die "Error closing file: $!";

  set_properties()
  
	"set_properties" 方法可被用于设置通过"Excel::Writer::XLSX"模块创建的Excel文件的文档属性。
    当你使用Excel中的"办公按钮" ->"准备"->"属性"选项时，可以看到这些属性。

	属性值应该以哈希格式传递，如下：

        $workbook->set_properties(
            title    => 'This is an example spreadsheet',
            author   => 'John McNamara',
            comments => 'Created with Perl and Excel::Writer::XLSX',
        );

    可以被设置的属性是:

        title    #标题
        subject	 #主题	
        author	 #作者
        manager  #经理
        company  #公司
        category #类别
        keywords #关键字
        comments #注释 
        status	 #状态

    请查看"properties.pl" 程序。

  define_name()
    
    该方法被用于定义一个名字，它能被用于表示工作簿中的一个值，一个单独的单元格，或一定范围内的单元格
	例如：设置一个 global/workbook 名:

        # Global/workbook names.
        $workbook->define_name( 'Exchange_rate', '=0.96' );
        $workbook->define_name( 'Sales',         '=Sheet1!$G$1:$H$10' );

    也可以使用语法"sheetname!definedname"在名字之前加上表名来定义一个 local/worksheet:
        # Local/worksheet name.
        $workbook->define_name( 'Sheet2!Sales',  '=Sheet2!$G$1:$G$10' );

    如果工作表名含有空格或特殊字符，你必须像在Excel中一样，用单引号将名字括起来：
        $workbook->define_name( "'New Data'!Sales",  '=Sheet2!$G$1:$G$10' );

    查看 defined_name.pl 程序。

  set_tempdir()
   
     "Excel::Writer::XLSX"在组装成最后的工作簿之前，把数据存储在临时文件中
     "File::Temp"模块用于创建这些临时文件。File::Temp模块使用"File::Spec"为这些临时文件指定一个合适的位置，例如"/tmp"或"c:\windows\temp".你可以按下面的方法找出你系统上哪个目录被使用了：
        perl -MFile::Spec -le "print File::Spec->tmpdir()
    如果默认的临时文件目录不能使用，你可以使用"set_tempdir()"方法指定一个可供选择的位置：
        $workbook->set_tempdir( '/tmp/writeexcel' );
        $workbook->set_tempdir( 'c:\windows\temp\writeexcel' );

    用于存放临时文件的目录必须先存在，“set_temp()”方法不会新建一个目录。
    一个潜在问题是一些Windows系统将并发临时文件的数量限制为大约800个。这意味着，一个在该种系统上运行的单个程序将会被限制创建总共800个工作簿和工作表对象。如果必要，你可以运行多个非并发程序来避免这种情况。
  set_custom_color( $index, $red, $green, $blue )
   #设置自定义颜色值
    "set_custom_color()"方法能用于使用更合适的颜色重载其中之一的内建颜色值。
	$index的值应该在8..63之间，查看see "COLOURS IN EXCEL".

	默认的命名颜色使用如下索引：

         8   =>   black
         9   =>   white
        10   =>   red
        11   =>   lime       #绿黄色
        12   =>   blue
        13   =>   yellow
        14   =>   magenta    #洋红色
        15   =>   cyan       #蓝绿色
        16   =>   brown
        17   =>   green
        18   =>   navy      #淡蓝色
        20   =>   purple    #紫色
        22   =>   silver    #银色
        23   =>   gray      #灰色
        33   =>   pink      #粉红色
        53   =>   orange

    使用它的RGB(red green blue)成分设置新颜色。 $red,$green 和 $blue的值范围必须在0..255之间。
	你可以在Excel中使用"工具"->选项->颜色->修改"对话框决定需要的颜色。
	
	"set_custom_color()"工作簿方法可以使用HTML风格的十六进制值：

        $workbook->set_custom_color( 40, 255,  102,  0 );       # Orange
        $workbook->set_custom_color( 40, 0xFF, 0x66, 0x00 );    # Same thing
        $workbook->set_custom_color( 40, '#FF6600' );           # Same thing

        my $font = $workbook->add_format( color => 40 );        # Modified colour

   
    "set_custom_color()"方法的返回值是被修改的颜色的索引：
        my $ferrari = $workbook->set_custom_color( 40, 216, 12, 12 );

        my $format = $workbook->add_format(
            bg_color => $ferrari,
            pattern  => 1,
            border   => 1
        );

    注意，在XLSX格式中，颜色调色板不确切局限为53种纯色。Excel::Writer::XLSX模块会在以后的阶段扩展以支持新的，半无限的调色板。
  sheets( 0, 1, ... )
   
	"sheets()"方法返回一个工作簿中工作表的列表或者列表切片

	如果没有传递参数给sheet()方法，则返回工作簿中的所有工作表。如果你想对一个工作表进行重复操作，这将很有用。

        for $worksheet ( $workbook->sheets() ) {
            print $worksheet->get_name();
        }

  
	你可以指定一个列表切片返回一个或多个工作表对象：
	
        $worksheet = $workbook->sheets( 0 );
        $worksheet->write( 'A1', 'Hello' );

    或者因为"sheets()"的返回值是一个对工作表对象的引用，你可以将上面的例子写为：
        $workbook->sheets( 0 )->write( 'A1', 'Hello' );

	
	下面的例子返回一个工作簿中的第一个和最后一个工作表：

        for $worksheet ( $workbook->sheets( 0, -1 ) ) {
            # Do something
        }


  set_1904()
   Excel将数据存储为实数，其整数部分存储自新纪元以来的天数，其小数部分存储一天的百分比。新纪元可以是1900或1904。Windows上的Excel使用1900，Mac上的Excel使用1904.然而，任何平台上的Excel都会在系统之间自动转换。
   Excel::Writer::XLSX默认使用1900格式存储数据。如果你想改变它，你可以调用"set_1904()"工作簿方法。对于1900它返回0，对于1904它返回1.


  set_optimization()
  
	"set_optimization()" 方法用于打开Excel::Writer::XLSX模块中的优化方案。目前只有一条减少内存使用的优化方案。

        $workbook->set_optimization();


    注意，打开此优化方案后，当通过"write_*()"方法中的其中之一在新行中添加一个单元格后，一列数据被写入然后被删除。因为一旦优化开启后，这样的数据应该以连续的行顺序写入。？
    
	该方法必须在任何调用"add_worksheet()"方法之前被调用。

WORKSHEET METHODS 工作表方法
	
	通过调用工作簿对象中的"add_worksheet()"方法创建一个新的工作表:

        $worksheet1 = $workbook->add_worksheet();
        $worksheet2 = $workbook->add_worksheet();

	下面的方法对于一个新的worksheet是可用的：

        write()
        write_number()
        write_string()
        write_rich_string()
        keep_leading_zeros()    #保留前导0
        write_blank()
        write_row()
        write_col()
        write_date_time()
        write_url()              #写入url
        write_url_range()
        write_formula()#写入公式
        write_comment()#写入注释
        show_comments()#显式注释
        set_comments_author()
        add_write_handler()
        insert_image()#插入图像
        insert_chart()#插入图表
        data_validation()#数据检验
        conditional_format()
        get_name()
        activate()#激活
        select()
        hide()
        set_first_sheet()
        protect()
        set_selection()
        set_row()
        set_column()
        outline_settings()
        freeze_panes()          #冻结窗格
        split_panes()		    #分割窗格
        merge_range()			#合并值域
        merge_range_type()
        set_zoom()
        right_to_left()
        hide_zero()				#隐藏0
        set_tab_color()			#设置标记颜色
        autofilter()			#自动筛选
        filter_column()
        filter_column_list()

  Cell notation 单元格表示法(先列-后行) 
  	
	Excel::Writer::XLSX支持两种形式的表示法来指定单元格的位置:行-列表示法和A1表示法。

    Row-column notation uses a zero based index for both row and column
    while A1 notation uses the standard Excel alphanumeric sequence of
    column letter and 1-based row. 例如，:

	行列表示法对行-列都使用以0为基础的索引
	而A1表示法使用标准的Excel字母数字序列为列，以1为基础作为行。例如：
	
        (0, 0)      # 最左最顶部的单元格（使用行-列表示法）
        ('A1')      # The top left cell in A1 notation.


        (1999, 29)  # 行-列表示法.
        ('AD2000')  # 使用A1表示法的同一单元格
#       单元格列的范围在Excel2003中是A..IV
  
	如果你提及单元格编程，行-列表示法很有用：

        for my $i ( 0 .. 9 ) {
            $worksheet->write( $i, 0, 'Hello' );    # Cells A1 to A10
        }

   
    A1表示法对于手动设置工作表和使用公式工作很有帮助：
        $worksheet->write( 'H1', 200 );
        $worksheet->write( 'H2', '=H1+1' ); #使用公式
Ecxel形如：
 ABCDEFGHIJKLMN
1
2
3
4
5
6
     在公式和可用的方法中你也可以使用"A:A"的列表示法：
	
        $worksheet->write( 'A1', '=SUM(B:B)' );

  	包含在套件中的Excel::Writer::XLSL::Utility 模块含有A1表示法的帮助函数，例如：
	
        use Excel::Writer::XLSX::Utility;

        ( $row, $col ) = xl_cell_to_rowcol( 'C2' );    # (1, 2)
        $str           = xl_rowcol_to_cell( 1, 2 );    # C2

  	简单地，在下面给出的章节中工作表方法调用的参数列表依据行-列表示法，任何情况下，都可以使用A1表示法
	
  
	注意：在Excel中也可以使用R1C1表示法。但Excel::Writer::XLSX不支持这。
	
  write( $row, $column, $token, $format )
 
	Excel的数据类型之间有区别，比如字符串，数字，空格，公式和超链接。为了简化写入数据的处理，write()方法为更多的特定方法指定一种普遍的别名：

        write_string()
        write_number()
        write_blank()
        write_formula()
        write_url()
        write_row()
        write_col()

	一般规则就是：如果数据看起来像什么那就写入什么。下面是用 行-列表示法和A1表示法写的例子：
                                                            # Same as:
        $worksheet->write( 0, 0, 'Hello'                 ); # write_string()
        $worksheet->write( 1, 0, 'One'                   ); # write_string()
        $worksheet->write( 2, 0,  2                      ); # write_number()
        $worksheet->write( 3, 0,  3.00001                ); # write_number()
        $worksheet->write( 4, 0,  ""                     ); # write_blank()
        $worksheet->write( 5, 0,  ''                     ); # write_blank()
        $worksheet->write( 6, 0,  undef                  ); # write_blank()
        $worksheet->write( 7, 0                          ); # write_blank()
        $worksheet->write( 8, 0,  'http://www.perl.com/' ); # write_url()
        $worksheet->write( 'A9',  'ftp://ftp.cpan.org/'  ); # write_url()
        $worksheet->write( 'A10', 'internal:Sheet1!A1'   ); # write_url()
        $worksheet->write( 'A11', 'external:c:\foo.xlsx' ); # write_url()
        $worksheet->write( 'A12', '=A3 + 3*A4'           ); # write_formula()
        $worksheet->write( 'A13', '=SIN(PI()/4)'         ); # write_formula()
        $worksheet->write( 'A14', \@array                ); # write_row()
        $worksheet->write( 'A15', [\@array]              ); # write_col()

        #如果设置了保留前置0属性：
		
        $worksheet->write( 'A16', 2                      ); # write_number()
        $worksheet->write( 'A17', 02                     ); # write_string()
        $worksheet->write( 'A18', 00002                  ); # write_string()

        # Write an array formula. Not available in Spreadsheet::WriteExcel.
         $worksheet->write( 'A19', '{=SUM(A1:B1*A2:B2)}'  ); # write_formula()

   
	"看起来像"的规则由正则表达式定义：
    "write_number()" 如果 $token 是一个基于如下正则的数字：
    "$token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/".

    "write_string()" 如果设置了"保留前导0"("keep_leading_zeros()")并且 $token 是一个基于如下正则的带有前导0的整数："$token =~/^0\d+$/".

    "write_blank()" 如果 $token 未定义或是一个空字符串: "undef", "" 或 ''.

    "write_url()" 如果 $token 是一个基于如下正则的 http, https, ftp 或者 mailto URL 
	              "$token =~ m|^[fh]tt?ps?://|" or "$token =~m|^mailto:|".

    "write_url()" 如果 $token 是一个基于如下正则的内部的或外部的引用：
                	"$token =~ m[^(in|ex)ternal:]".

    "write_formula()" 如果$token的第一个字符是"=".

    "write_array_formula()" 如果 $token 匹配 "/^{=.*}$/".

    "write_row()" 如果 $token 是一个数组引用.

    "write_col()" 如果 $token 是数组引用中的数组引用。

    "write_string()" 如果前面任一情况都不适用.

     $format 参数是可选的。它应该是个合法的格式对象。查看 "CELL FORMATTING":

        my $format = $workbook->add_format();
        $format->set_bold();
        $format->set_color( 'red' );
        $format->set_align( 'center' );

        $worksheet->write( 4, 0, 'Hello', $format );    # Formatted string

    write()方法会忽略空字符串或"undef"，除非提供了格式。就这点而论，你不必担心对空值或未定义值的处理。查看"write_blank()" 方法。
     "write()"方法的一个问题是，偶尔，数据看起来像一个数但是你不想把它看作一个数字。例如，邮政编码或ID号常以前导0开头。如果将该数据作为数字写入，则前导0会被删除。你可以使用"keep_leading_zeros()"方法改变该默认行为。当这个特性起作用时，任何带有前导0的整数会被当作字符串并且前导0被保留。查看"keep_leading_zeros()"章节获取该问题的详细信息。
    
	你也可以使用"add_write_handler()"把你自己的数据处理器添加到"write()"方法。

    "write()"方法也会处理UTF-8格式的Unicode字符串。
    "write" 方法返回:

        0 成功.
       -1 参数个数不足
       -2 行或列超限
       -3 字符串过长

  write_number( $row, $column, $number, $format )
    向行和列指定（$row and $column）的单元格中写入整数或浮点数。
        $worksheet->write_number( 0, 0, 123456 );
        $worksheet->write_number( 'A2', 2.3451 );

     $format 参数可选.

    一般地，使用"write()"方法就足够了。
    注意：有些版本的Excel2007不显示由Excel::Writer::XLSX写入的公式计算值。求助于Excel的所有可用服务包来修复该问题。
  write_string( $row, $column, $string, $format )
   
    向行和列指定的单元格中写入字符串：
        $worksheet->write_string( 0, 0, 'Your text here' );
        $worksheet->write_string( 'A2', 'or here' );

   最大的字符串长度为32767个字符。然而，Excel单元格中能显示的最大字段是1000个。所有的32767个字符可以显示在一个公式栏中。
    $format参数是可选的.

   
    "write()" 方法也会处理UTF-8格式的字符串。请查看"unicode_*.pl"程序
   
    一般地，使用"write()" 方法就足够了。	然而，你有时候可能会使用"write_string()"方法去写入看起来像数字但你又不想把它看作数字的数据。例如，邮政编码或电话号码：
        # 作为普通的字符串写入
        $worksheet->write_string( 'A1', '01209' );

   然而，如果用户编辑该字符串，Excel可能会把字符串转换回数字。你可以使用Excel的文本格式"@"来避免它:
        # 格式化为字符串.编辑时不转换为数字。
        my $format1 = $workbook->add_format( num_format => '@' );
        $worksheet->write_string( 'A2', '01209', $format1 );

    write_rich_string( $row, $column, $format, $string, ..., $cell_format )

    "write_rich_string()"方法用于写入带有多种格式的字符串。例如，写入字符串"This is bold and this is italic" 你可以使用下面的方法： 
        my $bold   = $workbook->add_format( bold   => 1 );
        my $italic = $workbook->add_format( italic => 1 );

        $worksheet->write_rich_string( 'A1',
            'This is ', $bold, 'bold', ' and this is ', $italic, 'italic' );

    
    基本规则是把字符串分段并把$format格式对象放在你想格式化的片段之前。例如：
        # 未格式化的字符串
          'This is an example string'

        # 分割
          'This is an ', 'example', ' string'

        # 在你想格式化的片段前添加格式
          'This is an ', $format, 'example', ' string'

        # In Excel::Writer::XLSX.
        $worksheet->write_rich_string( 'A1',
            'This is an ', $format, 'example', ' string' );

     没有格式的字符串片段使用默认的格式。例如，当写入字符串"Some bold text"你会使用下面的第一个例子，但是它与第二个例子等价。
        # 使用默认格式:
        my $bold    = $workbook->add_format( bold => 1 );

        $worksheet->write_rich_string( 'A1',
            'Some ', $bold, 'bold', ' text' );

        # 或更明确地:
        my $bold    = $workbook->add_format( bold => 1 );
        my $default = $workbook->add_format();

        $worksheet->write_rich_string( 'A1',
            $default, 'Some ', $bold, 'bold', $default, ' text' );

    对于Excel，只有格式的字体属性诸如字体名，风格，大小，下划线，颜色和效果被应用到字符串片段上。其他属性诸如边框，背景，对齐方式必须被应用于单元格。
	
    "write_rich_string()"方法允许你把最后一个参数作为单元格格式使用（如果它是一个格式对象的话）来完成以上功能。下面的例子是使单元格中的rich string 居中对齐。

        my $bold   = $workbook->add_format( bold  => 1 );
        my $center = $workbook->add_format( align => 'center' );

        $worksheet->write_rich_string( 'A5',
            'Some ', $bold, 'bold text', ' centered', $center );

    查看"rich_strings.pl" 获取详细信息

        my $bold   = $workbook->add_format( bold        => 1 );
        my $italic = $workbook->add_format( italic      => 1 );
        my $red    = $workbook->add_format( color       => 'red' );
        my $blue   = $workbook->add_format( color       => 'blue' );
        my $center = $workbook->add_format( align       => 'center' );
        my $super  = $workbook->add_format( font_script => 1 );


        # 使用多种格式写入一些字符串
        $worksheet->write_rich_string( 'A1',
            'This is ', $bold, 'bold', ' and this is ', $italic, 'italic' );

        $worksheet->write_rich_string( 'A3',
            'This is ', $red, 'red', ' and this is ', $blue, 'blue' );

        $worksheet->write_rich_string( 'A5',
            'Some ', $bold, 'bold text', ' centered', $center );

        $worksheet->write_rich_string( 'A7',
            $italic, 'j = k', $super, '(n-1)', $center );

    正如 "write_sting()" 一样，它可写入的最大字符数是 32767个. 

  keep_leading_zeros()

    当使用"write()"方法时， keep_leading_zeros()方法改变了带有前导0整数的默认处理方式。
     "write()"方法使用正则表达式来决定写入什么样的数据到Excel工作表中。如果数据看起来像数字它就使用"write_number()"方法写入数字。该方法的一个问题是偶尔数据看起来像数字但你不想将它看作一个数字。
	
   	例如邮政编码和ID号，常以前导0开头。如果你把这样的数据当作数字写入，则前导0被删除。当你手动在Excel中输入数据时，这也是默认行为。

     为了避免此问题，你可以使用三选项之一。写入一个格式化后的数字、将数字当作字符串写入或使用"keep_leading_zeros()"方法来改变"write()"方法的默认行为：
        # 隐式地写入一个数字,前导0被删除: 1209
        $worksheet->write( 'A1', '01209' );

        #使用格式写入以0填充的数字: 01209
        my $format1 = $workbook->add_format( num_format => '00000' );
        $worksheet->write( 'A2', '01209', $format1 );

        # 显式地当作字符串写入: 01209
        $worksheet->write_string( 'A3', '01209' );

        # 隐式地当作字符串写入: 01209
        $worksheet->keep_leading_zeros();
        $worksheet->write( 'A4', '01209' );

  
	上面的代码会生成一个如下所示的工作表:

         -----------------------------------------------------------
        |   |     A     |     B     |     C     |     D     | ...
         -----------------------------------------------------------
        | 1 |      1209 |           |           |           | ...
        | 2 |     01209 |           |           |           | ...
        | 3 | 01209     |           |           |           | ...
        | 4 | 01209     |           |           |           | ...

   
    例子里单元格在不同的边上是因为Excel默认以左对齐方式显式字符串，以右对齐方式显式数字。

    应该注意的是如果用户编辑例子中的"A3"和"A4"数据，字符串会恢复为数字。这还是Excel的默认行为。使用文本格式"@"可以避免该行为：
        # Format as a string (01209)
        my $format2 = $workbook->add_format( num_format => '@' );
        $worksheet->write_string( 'A5', '01209', $format2 );

   
    "keep_leading_zeros()"特性默认是关闭的，它以0或1为参数。如果没有给它指定参数，默认为1：
        $worksheet->keep_leading_zeros(   )     # Set on
        $worksheet->keep_leading_zeros( 1 );    # Set on
        $worksheet->keep_leading_zeros( 0 );    # Set off

   
  write_blank( $row, $column, $format )
   
    写入由行和列指定的空白单元格：
        $worksheet->write_blank( 0, 0, $format );
    
    该方法用于向不含字符串或数字值的单元格添加格式。
    Excel中"空"单元格和"空白"单元格不同。空单元格不包含数据，空白单元格不包含数据但是却含有格式。Excel存储"空白"单元格但忽略“空”单元格。
    
	像这样，如果你写入的单元格为空且不含格式，它会被忽略：
        $worksheet->write( 'A1', undef, $format );    # write_blank()
        $worksheet->write( 'A2', undef );             # Ignored

	这看起来很乏味的事实意味着你可以不用特别处理"undef"和空字符串值写入数组的数据。
    

  write_row( $row, $column, $array_ref, $format )
   "write_row()"方法能用于将1维数组或2维数组的数据合为一体。这对于将数据库查询结果转换为Excel工作表很有用。 你必须传递给数组数据一个引用而不是数组本身。"write()" 方法被数据中的每个元素调用，例如：
    
        @array = ( 'awk', 'gawk', 'mawk' );
        $array_ref = \@array;

        $worksheet->write_row( 0, 0, $array_ref );

        # The above example is equivalent to:
        $worksheet->write( 0, 0, $array[0] );
        $worksheet->write( 0, 1, $array[1] );
        $worksheet->write( 0, 2, $array[2] );

    注意：为方便起见，如果传递的参数是数组引用，则"write()"方法和"write_row()"的行为一样。因此，下面的2种方法调用等价：
        $worksheet->write_row( 'A1', $array_ref );    # Write a row of data
        $worksheet->write(     'A1', $array_ref );    # Same thing

    正如所有的write方法一样，$format参数是可选的.如果指定了一个格式，它会被应用到数据数组的所有元素上。 

   
    数据中的数组引用会被当作列。这允许你一举写入2维数组的数据，例如：
        @eec =  (
                    ['maggie', 'milly', 'molly', 'may'  ],
                    [13,       14,      15,      16     ],
                    ['shell',  'star',  'crab',  'stone']
                );

        $worksheet->write_row( 'A1', \@eec );

	会产生如下的工作表：
         -----------------------------------------------------------
        |   |    A    |    B    |    C    |    D    |    E    | ...
         -----------------------------------------------------------
        | 1 | maggie  | 13      | shell   | ...     |  ...    | ...
        | 2 | milly   | 14      | star    | ...     |  ...    | ...
        | 3 | molly   | 15      | crab    | ...     |  ...    | ...
        | 4 | may     | 16      | stone   | ...     |  ...    | ...
        | 5 | ...     | ...     | ...     | ...     |  ...    | ...
        | 6 | ...     | ...     | ...     | ...     |  ...    | ...

  
    以行-列顺序写入数据，请参考下面的 "write_col()"方法。
   
    除非对数据应用一种格式，否则数据中的任何"未定义"值将被忽略，这种情况下，一个格式化后的空白单元格会被写入。2种情况下，适当的行和列的值仍然会增加。
   
    
    当写入数据的元素时，"write_row()" 方法返回出现的第一个错误，如果没有出现错误，返回0。
    请查看 "write_arrays.pl" 程序。


    "write_row()"方法在文本文件到Excel文件之间允许如下惯用转换：
        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        my $workbook  = Excel::Writer::XLSX->new( 'file.xlsx' );
        my $worksheet = $workbook->add_worksheet();

        open INPUT, 'file.txt' or die "Couldn't open file: $!";

        $worksheet->write( $. -1, 0, [split] ) while <INPUT>;

  write_col( $row, $column, $array_ref, $format )
     "write_col()"方法能用于一次性地写入一维或二维数组的数据。这对于将数据库查询结果转换为Excel工作表很有用。你必须传递一个数组的引用而非数组本身。 "write()" 方法随后被数据的每个元素调用。 例如，
	 
    "write_col()"方法能被用于一举写入1维或2维数组的数据。这对于将数据库查询结果转换为Excel工作表很有用。你必须传递给数组数据一个引用而不是数组本身。"write()" 方法随后被数据中的每个元素调用，例如：
        @array = ( 'awk', 'gawk', 'mawk' );
        $array_ref = \@array;

        $worksheet->write_col( 0, 0, $array_ref );

        # 上面的例子等价于:
        $worksheet->write( 0, 0, $array[0] );
        $worksheet->write( 1, 0, $array[1] );
        $worksheet->write( 2, 0, $array[2] );

    正如所有的write方法一样，$format参数是可选的.如果指定了格式，它会被应用到数组所有的元素上。 


    数据中的数组引用会被当作行来对待。这允许你一举写入2维数组的数据。例如：
        @eec =  (
                    ['maggie', 'milly', 'molly', 'may'  ],
                    [13,       14,      15,      16     ],
                    ['shell',  'star',  'crab',  'stone']
                );

        $worksheet->write_col( 'A1', \@eec );

    会产生如下的工作表：

         -----------------------------------------------------------
        |   |    A    |    B    |    C    |    D    |    E    | ...
         -----------------------------------------------------------
        | 1 | maggie  | milly   | molly   | may     |  ...    | ...
        | 2 | 13      | 14      | 15      | 16      |  ...    | ...
        | 3 | shell   | star    | crab    | stone   |  ...    | ...
        | 4 | ...     | ...     | ...     | ...     |  ...    | ...
        | 5 | ...     | ...     | ...     | ...     |  ...    | ...
        | 6 | ...     | ...     | ...     | ...     |  ...    | ...

 
	将数据以列-行的顺序写入请查看上面的"write_row()"方法。

    除非对数据应用一种格式，否则数据中的任何"未定义"值将被忽略，这种情况下，一个格式化后的空白单元格会被写入。2种情况下，适当的行和列的值仍然会增加。


	就如上面所说的，"write()" 方法能被用作 "write_row()" 的同义词，并且 "write_row()"方法将嵌套的数组引用当作列.
   
    所以，下面的2个方法调用是等价的，尽管对"write_col()"更清楚的调用可维护性更好。
        $worksheet->write_col( 'A1', $array_ref     ); # Write a column of data
        $worksheet->write(     'A1', [ $array_ref ] ); # Same thing

   
    当写入数据的元素时，"write_col()" 方法返回出现的第一个错误，如果没有出现错误，返回0。
    查看上面的“write”方法描述的返回值。 
    
    也请查看 "write_arrays.pl"程序。

  write_date_time( $row, $col, $date_string, $format )
   
    "write_date_time()" 方法能用于向指定行列的单元格里写入日期或时间：
        $worksheet->write_date_time( 'A1', '2004-05-13T23:20', $date_format );

    $date_string 应该是下列的格式:

        yyyy-mm-ddThh:mm:ss.sss

    它遵守ISO8601日期但是应该注意的是不支持全部范围内的ISO8601格式。
    下面的变更的$date_string 参数是允许的：
        yyyy-mm-ddThh:mm:ss.sss         # Standard format
        yyyy-mm-ddT                     # No time
                  Thh:mm:ss.sss         # No date
        yyyy-mm-ddThh:mm:ss.sssZ        # Additional Z (but not time zones)
        yyyy-mm-ddThh:mm:ss             # No fractional seconds
        yyyy-mm-ddThh:mm                # No seconds

    注意"T"在所有情况下是必须的。
   
    日期应该一直有个$format格式，否则他会以数字形式出现。这里是典型的例子：
        my $date_format = $workbook->add_format( num_format => 'mm/dd/yy' );
        $worksheet->write_date_time( 'A1', '2004-05-13T23:20', $date_format );

     对于1900纪元，合法的日期应该在1900-01-01到9999-12-31之间，而对于1904纪元，合法的日期是1904-01-01 到 9999-12-31。对于Excel，超出这些范围的日期会被当作字符串写入。
    请查看 date_time.pl程序。

  write_url( $row, $col, $url, $format, $label )
     将URL的超链接写入由行和列指定的单元格中。超链接由2部分组成：可见的标记和不可见的链接。可见的标记与链接一样，除非指定了可替换的标记。$label参数是可选的。标记由"write()"方法写入。因此，可以将字符串、数字或公式作为标记。
    
	$format参数也是可选的，然而，没有格式的话，链接就不像一个格式了。

    建议的格式是:

        my $format = $workbook->add_format( color => 'blue', underline => 1 );

    注意，如果用户没有指定一种格式，该行为与Spreadsheet::WriteExcel提供的默认的超链接格式不同。

	支持4种web风格的URI "http://", "https://", "ftp://" 和 "mailto:":

        $worksheet->write_url( 0, 0, 'ftp://www.perl.org/', $format );
        $worksheet->write_url( 1, 0, 'http://www.perl.com/', $format, 'Perl' );
        $worksheet->write_url( 'A3', 'http://www.perl.com/',      $format );
        $worksheet->write_url( 'A4', 'mailto:jmcnamara@cpan.org', $format );

    There are two local URIs supported: "internal:" and "external:". These
    are used for hyperlinks to internal worksheet references or external
    workbook and worksheet references:
    支持2种本地URLs："internal:" 和 "external:。这些用于内部工作表引用的超链接或外部工作簿或工作表的引用：
        $worksheet->write_url( 'A6',  'internal:Sheet2!A1',              $format );
        $worksheet->write_url( 'A7',  'internal:Sheet2!A1',              $format );
        $worksheet->write_url( 'A8',  'internal:Sheet2!A1:B2',           $format );
        $worksheet->write_url( 'A9',  q{internal:'Sales Data'!A1},       $format );
        $worksheet->write_url( 'A10', 'external:c:\temp\foo.xlsx',       $format );
        $worksheet->write_url( 'A11', 'external:c:\foo.xlsx#Sheet2!A1',  $format );
        $worksheet->write_url( 'A12', 'external:..\foo.xlsx',            $format );
        $worksheet->write_url( 'A13', 'external:..\foo.xlsx#Sheet2!A1',  $format );
        $worksheet->write_url( 'A13', 'external:\\\\NET\share\foo.xlsx', $format );

  
	典型的工作表引用形式是"Sheet1!A1"。你可以使用标准的Excel表示法"Sheet1!A1:B2"来指定工作表范围。

    In external links the workbook and worksheet name must be separated by
    the "#" character: "external:Workbook.xlsx#Sheet1!A1'".
	外部链接中，工作簿和工作表的名字必须用"#"隔开："external:Workbook.xlsx#Sheet1!A1'"

    You can also link to a named range in the target worksheet. 例如，
    say you have a named range called "my_name" in the workbook
    "c:\temp\foo.xlsx" you could link to it as follows:
    你也可以链接到目标工作表中的一个命名值域上。例如，比如在"c:\temp\foo.xlsx"工作簿中你有一个具名的值域叫做"my_name"，你可以按下面的方法链接到它。
        $worksheet->write_url( 'A14', 'external:c:\temp\foo.xlsx#my_name' );

    Excel要求包含空格或非字母字符的工作表名要用单引号引起来，如下"'Sales Data'!A1"。如果你使用单引号引起的字符串作工作表名，你需要使用\'转义单引号，或使用q{}。
  
    也支持到网络文件的链接。MS/Novell 网络文件一般以2个反斜杠开头，形如"\\NETWORK\etc"。为了在单引号或双引号字符串中生成这样的网络名，需要转义反斜杠，'\\\\NETWORK\etc'.

    如果你使用双引号字符串，你应该注意转义任何看起来像元字符的字符。更多信息，查看Perlfaq5：为什么不能在DOS路径中使用 "C:\temp\foo" in DOS paths?"。
  
  
    最后，你可以使用左斜杠来避免大部分引号问题。左斜杠在内部被转换成反斜杠：
        $worksheet->write_url( 'A14', "external:c:/temp/foo.xlsx" );
        $worksheet->write_url( 'A15', 'external://NETWORK/share/foo.xlsx' );

    

  write_formula( $row, $column, $formula, $format, $value )
	将公式或函数写入由行和列指定的单元格中。
        $worksheet->write_formula( 0, 0, '=$B$3 + B4' );
        $worksheet->write_formula( 1, 0, '=SIN(PI()/4)' );
        $worksheet->write_formula( 2, 0, '=SUM(B1:B5)' );
        $worksheet->write_formula( 'A4', '=IF(A3>1,"Yes", "No")' );
        $worksheet->write_formula( 'A5', '=AVERAGE(1, 2, 3, 4)' );
        $worksheet->write_formula( 'A6', '=DATEVALUE("1-Jan-2001")' );

    同样也支持数组公式:

        $worksheet->write_formula( 'A7', '{=SUM(A1:B1*A2:B2)}' );

 
    如果需要，可以指定公式的计算结果。当与不能计算公式值的非Excel应用程序一起工作时，这偶尔需要。计算过的$value值被添加在参数列表的末尾：
        $worksheet->write( 'A1', '=2+2', $format, 4 );


  write_array_formula($first_row, $first_col, $last_row, $last_col, $formula, $format, $value)
     将数组公式写入到一个单元格值域中。在Excel中一个数组公式就是在一组值上执行计算的公式。它可以返回单个值或一个值域。
    公式两边的一对花括号表明这是一个数组公式："{=SUM(A1:B1*A2:B2)}"。如果数组公式返回单个值，则$first_ 和 $last_ 参数应该一样：
        $worksheet->write_array_formula('A1:A1', '{=SUM(B1:C1*B2:C2)}');


	此种情况下仅仅使用"write_formula()"或 "write()"方法更容易些：

        # 和上面的一样，但是更简洁：
        $worksheet->write( 'A1', '{=SUM(B1:C1*B2:C2)}' );
        $worksheet->write_formula( 'A1', '{=SUM(B1:C1*B2:C2)}' );

    For array formulas that return a range of values you must specify the
    range that the return values will be written to:

        $worksheet->write_array_formula( 'A1:A3',    '{=TREND(C1:C3,B1:B3)}' );
        $worksheet->write_array_formula( 0, 0, 2, 0, '{=TREND(C1:C3,B1:B3)}' );

   如果需要，可以指定公式的计算结果。当与不能计算公式值的非Excel应用程序一起工作时，这偶尔需要。计算过的$value值被添加在参数列表的末尾：

        $worksheet->write_array_formula( 'A1:A3', '{=TREND(C1:C3,B1:B3)}', $format, 105 );

     此外，一些Excel2007的早期版本没有提供数组公式，所以它们不计算数组公式的值。安装最新的Office Service Pack可以修复该问题。
    查看 "array_formula.pl" 程序。

   注意：Spreadsheet::WriteExcel不支持数组公式。
  store_formula( $formula )
   
    不宜使用。这是一个Spreadsheet::WriteExcel 方法，Excel::Writer::XLSX已经不在需要了。看下面。
  repeat_formula( $row, $col, $formula, $format )
 
     不宜使用。这是一个Spreadsheet::WriteExcel 方法，Excel::Writer::XLSX.已经不在需要了。

    在 Spreadsheet::WriteExcel中写入公式其计算是昂贵的，因为它们由递归下降的解释器解析。
	"store_formula()" 和 "repeat_formula()" 方法再利用预解析公式作为避免重复公式的系统开销的方法之一。
   
    在Excel::Writer::XLSX中，这不再需要，因为写入公式就像写入字符串和数字一样快。
    The methods remain for backward compatibility but new
    Excel::Writer::XLSX programs shouldn't use them.
    以上方法向后兼容，但是新的Excel::Writer::XLSX 不能使用它们。
  write_comment( $row, $column, $string, ... )
  
    "write_comment()" 方法用于向单元格中写入注释。
	单元格中的注释由Excel单元格右上角的小红色三角标注。将光标移到红色三角上将显示注释。 
    下面的例子显式了如何在单元格中添加注释：
        $worksheet->write        ( 2, 2, 'Hello' );
        $worksheet->write_comment( 2, 2, 'This is a comment.' );


	通常，你可以用一个"A1"单元格引用代替 $row 和 $column 参数：

        $worksheet->write        ( 'C3', 'Hello');
        $worksheet->write_comment( 'C3', 'This is a comment.' );

     "write_comment()" 方法也处理UTF-8格式的字符串。

        $worksheet->write_comment( 'C3', "\x{263a}" );       # Smiley
        $worksheet->write_comment( 'C4', 'Comment ca va?' );

 
    除了基本的3个参数形式的"write_comment()"，你可以传递几对可选的 键值对来控制注释的格式：
        $worksheet->write_comment( 'C3', 'Hello', visible => 1, author => 'Perl' );

   
    这些选项的大多数很特别，并且通常默认的注释行为就是你想要的。然而，如果你需要对单元格注释更好的控制，那下面的选项可以使用：
        author
        visible
        x_scale
        width
        y_scale
        height
        color
        start_cell
        start_row
        start_col
        x_offset
        y_offset

    Option: author
    
        该选项用于表明谁是该单元格注释的作者。Excel在工作表底部的状态条栏中显式注释的作者。
            $worksheet->write_comment( 'C3', 'Atonement', author => 'Ian McEwan' );


        所有的单元格注释默认的作者能使用 "set_comments_author()"方法设置，看下面：
            $worksheet->set_comments_author( 'Perl' );

    Option: visible
    
        该选项能用于当打开工作表时，使单元格注释可见。Excel中默认的行为是注释被隐藏。然而，在Excel中也可以使单个注释或所有注释可见。在Excel::Writer::XLSX中，可以按下面的方法使单个注释可见：
            $worksheet->write_comment( 'C3', 'Hello', visible => 1 );

      
        使用"show_comments()" 工作表方法（看下面）可使工作表中的所有注释可见。
		相应地，如果所有的单元格注释都可见了，你可以隐藏单个注释：
            $worksheet->write_comment( 'C3', 'Hello', visible => 0 );

    Option: x_scale
        This option is used to set the width of the cell comment box as a
        factor of the default width.
        该选项用于设置单元格注释框的宽度：
            $worksheet->write_comment( 'C3', 'Hello', x_scale => 2 );
            $worksheet->write_comment( 'C4', 'Hello', x_scale => 4.2 );

    Option: width
     
        该选项用于设置单元格注释框的宽度，以像素表示
            $worksheet->write_comment( 'C3', 'Hello', width => 200 );

    Option: y_scale
        This option is used to set the height of the cell comment box as a
        factor of the default height.
        该选项用于设置单元格注释框的高度：
            $worksheet->write_comment( 'C3', 'Hello', y_scale => 2 );
            $worksheet->write_comment( 'C4', 'Hello', y_scale => 4.2 );

    Option: height
        This option is used to set the height of the cell comment box
        explicitly in pixels.
         该选项用于设置单元格注释框的高度，以像素表示
            $worksheet->write_comment( 'C3', 'Hello', height => 200 );

    Option: color

        该选项用于设置单元格注释框的背景色。你可以使用Excel::Writer::XLSX 可识别的具名颜色或颜色索引。
            $worksheet->write_comment( 'C3', 'Hello', color => 'green' );
            $worksheet->write_comment( 'C4', 'Hello', color => 0x35 );      # Orange

    Option: start_cell
   
        该选项用于设置注释将出现在哪个单元格。By default Excel displays comments one cell to the right and
        one cell above the cell to which the comment relates.
		然而，你可以改变它的默认行为，如果你愿意的话。在下面的例子中，默认会在单元格“D2”中出现的注释会移动到“E2”中。
            $worksheet->write_comment( 'C3', 'Hello', start_cell => 'E2' );
2011-12-5
    Option: start_row
    
        该选项用于设置注释将会出现在哪一行。行从0开始索引。
            $worksheet->write_comment( 'C3', 'Hello', start_row => 0 );

    Option: start_col
   
		该选项用于设置注释将会在哪一列出现。列从0开始索引。

            $worksheet->write_comment( 'C3', 'Hello', start_col => 4 );

    Option: x_offset
        This option is used to change the x offset, in pixels, of a comment
        within a cell:
        	该选项用于改变单元格中注释的x轴方向的偏移量，以像素计算。
            $worksheet->write_comment( 'C3', $comment, x_offset => 30 );

    Option: y_offset
        This option is used to change the y offset, in pixels, of a comment
        within a cell:
		该选项用于改变单元格中注释的y轴方向的偏移量，以像素计算。

            $worksheet->write_comment('C3', $comment, x_offset => 30);

    

	注意，使用诸如start_cell, start_row, start_col, x_offset 和 y_offset的选项调整单元格注释位置：Excel仅仅在单元格可见时才显示单元格注释的偏移量。当你把鼠标移到它们上面时，Excel不显示隐藏的单元格。

   
	注意行高和注释。如果你指定含有注释的单元格的行高，则Excel::Writer::XLSX会调整注释的高度以保持默认高度或用户指定的尺寸。然而，如果设置了文本框属性或在单元格中使用了很大的字体，Excel会自动调整行高。这意味着行高对于运行时的模块是未知的，因此注释框随着行一起扩展。使用 "set_row()"方法显式地指定行高来避免这个问题。

  show_comments()

    该方法用于当打开工作表时，让所有单元格注释可见
        $worksheet->show_comments();


    使用"write_comment"方法的"visible"参数能使单个的注释可见（看下面）：
        $worksheet->write_comment( 'C3', 'Hello', visible => 1 );

	如果所有的单元格注释都可见了，你可以按下面的方法隐藏单个注释：

        $worksheet->show_comments();
        $worksheet->write_comment( 'C3', 'Hello', visible => 0 );

  set_comments_author()
    该方法用于设置单元格注释的默认作者。
        $worksheet->set_comments_author( 'Perl' );

    使用"write_comment"方法的"author"参数设置单个注释的作者（看上面）。

	如果没有指定作者，默认的注释作者是空字符串,''.

  add_write_handler( $re, $code_ref )
   
	该方法用于扩展 Excel::Writer::XLSX 的"write()"方法来处理用户定义的数据。

	如果你查看上面章节的"write()"方法，你会发现它是几个更特殊的"write_*"方法的别名。然而，它并不总是如你所愿地正确。

 
    一种方法是你自己过滤输入数据并调用合适的 "write_*" 方法。
	另一种方法是使用"add_write_handler()"方法把你自己的自动化行为添加到"write()"方法中。
  
  
	"add_write_handler()" 有2个参数，$re,一个匹配输入数据的正则表达式；
	还有$code_ref，一个回调函数，来处理匹配后的数据：

        $worksheet->add_write_handler( qr/^\d\d\d\d$/, \&my_write );

    (在这些例子中"qr"操作符用于引起正则表达式字符串).


	该方法使用如下。假如你想写入7个数字的ID号作为字符串并保留任何前导0，你可以按下面的做：

        $worksheet->add_write_handler( qr/^\d{7}$/, \&write_my_id );


        sub write_my_id {
            my $worksheet = shift;
            return $worksheet->write_string( @_ );
        }

    * 你也可以使用"keep_leading_zeros()"方法.

	然后，如果你使用一个合适的字符串调用"write()"方法，它会被自动处理：
        # 写入 0000000.正常地，它会被写作数字0：
        $worksheet->write( 'A1', '0000000' );

	回调函数会接受一个对被调用工作表的引用，并且其它所有的参数被传递给"write()".回调函数会看到如下所示的@_参数列表：

        $_[0]   A ref to the calling worksheet. *
        $_[1]   Zero based row number.
        $_[2]   Zero based column number.
        $_[3]   A number or string or token.
        $_[4]   A format ref if any.
        $_[5]   Any other arguments.
        ...

        *  It is good style to shift this off the list so the @_ is the same
           as the argument list seen by write().


	你的回调函数应该使用"return()"返回"write_*" 方法的返回值，
	或返回“undef”来表明你拒绝了匹配并且想让"write()"方法正常继续。

    So 例如， if you wished to apply the previous filter only to ID
    values that occur in the first column you could modify your callback
    function as follows:
	所以，例如，如果你想把前面的过滤只应用到出现在第一列的ID值上，你可以按照下面修改你的回调函数：

        sub write_my_id {
            my $worksheet = shift;
            my $col       = $_[1];

            if ( $col == 0 ) {
                return $worksheet->write_string( @_ );
            }
            else {
                # Reject the match and return control to write()
                return undef;
            }
        }

    Now, you will get different behaviour for the first column and other
    columns:
    现在，你会在、第一列和其它列上得到不同的行为：
        $worksheet->write( 'A1', '0000000' );    # Writes 0000000
        $worksheet->write( 'B1', '0000000' );    # Writes 0

	你可以添加多个处理程序，此时它们会按照被添加的顺序调用。

    
	注意，"add_write_handler()"特别适合处理数据。

    查看 "write_handler 1-4" 程序。

  insert_image( $row, $col, $filename, $x, $y, $scale_x, $scale_y )
    
	部分支持。目前只对96像素的图像有效。这会在下次发布中修复。

   
    该方法用于向工作表中插入图像。图像格式可以是PNG, JPEG 或 BMP。 $x, $y, $scale_x 和 $scale_y 参数是可选的。
    	$worksheet1->insert_image( 'A1', 'perl.bmp' );
        $worksheet2->insert_image( 'A1', '../images/perl.bmp' );
        $worksheet3->insert_image( 'A1', '.c:\images\perl.bmp' );

    
	$x 和 $y参数可以用于指定到由$row和$col指定的单元格的左上角的偏移量。偏移量值以像素计算：

        $worksheet1->insert_image('A1', 'perl.bmp', 32, 10);

   	偏移量可以比图像下面的单元格的高度或宽度大。如果你想在同一个单元格中对齐一个或多个图像，这偶尔会有用。

   
	参数$scale_x 和 $scale_y能用于水平和垂直地测量插入的图片：

        # Scale the inserted image: width x 2.0, height x 0.8
        $worksheet->insert_image( 'A1', 'perl.bmp', 0, 0, 2, 0.8 );

    查看那"images.pl" 程序。


	注意：如果你想改变图像所占的任一行或列的默认尺寸，你必须在"insert_image()"之前调用"set_row()" 或 "set_column()"。
	如果你使用的字体比默认的大，行的高度也会改变。这反过来也会影响你图像的尺寸。你应该使用"set_row()"显式地设置行高来避免这个问题，如果它包含了会改变行高的字体大小。


	BMP图像应该是24比特，颜色为真彩，位图。通常，最好应避使用BMP图像，因为它们没有被压缩。

  insert_chart( $row, $col, $chart, $x, $y, $scale_x, $scale_y )
 	该方法用于向工作表中插入一个图表对象。图表必须由"add_chart()"工作表方法创建，并且它必须设置了 "embedded"选项。

        my $chart = $workbook->add_chart( type => 'line', embedded => 1 );

        # Configure the chart.
        ...

        # Insert the chart into the a worksheet.
        $worksheet->insert_chart( 'E2', $chart );


	查看"add_chart()" 获取关于怎样创建图表对象的细节，
	查看Excel::Writer::XLSX::Chart获取关于怎样配置图表的细节。查看"chart_*.pl"程序。

    $x, $y, $scale_x 和 $scale_y 参数是可选的 。
	


	$x 和 $y 能用于指定到由$row和$col指定的单元格左上角的偏移量，偏移量值以像素计算。

        $worksheet1->insert_chart( 'E2', $chart, 3, 3 );

    
	参数$scale_x 和 $scale_y能用于从水平方向和垂直方向测量图像：

        # Scale the width by 120% and the height by 150%
        $worksheet->insert_chart( 'E2', $chart, 0, 0, 1.2, 1.5 );

  data_validation()
    
	"data_validation()"方法用于构建Excel数据检验或限制用户输入到一个下拉列表中

        $worksheet->data_validation('B3',
            {
                validate => 'integer',   #验证
                criteria => '>',         #标准
                value    => 100,         #值
            });

        $worksheet->data_validation('B5:B9',
            {
                validate => 'list',
                value    => ['open', 'high', 'close'],
            });

 
	该方法包含很多参数，并在单独的章节“DATA VALIDATION IN EXCEL”中有详细描述。

	查看"data_validate.pl" 程序。
	
  conditional_format()
  
	"conditional_format()" 方法用于向单元格或基于用户自定义标准的单元格范围中添加格式

        $worksheet->conditional_formatting( 'A1:J10',
            {
                type     => 'cell',
                criteria => '>=',
                value    => 50,
                format   => $format1,
            }
        );


	该方法包含很多参数，并在单独的章节“CONDITIONAL FORMATTING IN EXCEL”中有详细描述

	查看"conditional_format.pl"程序。

  get_name()
   
	"get_name()" 方法用于检索工作表的名字。例如：

        for my $sheet ( $workbook->sheets() ) {
            print $sheet->get_name();
        }

	由于Excel::Writer::XLSX的设计和Excel的内部原因，没有设计"set_name()" 方法。
	设置工作表名字的唯一方法是通过"add_worksheet()"方法。

  activate()

	"activate()"方法用于指定在一个含有多个工作表的工作簿中，哪一个工作表是初始可见的：

        $worksheet1 = $workbook->add_worksheet( 'To' );
        $worksheet2 = $workbook->add_worksheet( 'the' );
        $worksheet3 = $workbook->add_worksheet( 'wind' );

        $worksheet3->activate();

   
	这与Excel VBA的active方法相似。可以通过"select()"方法选取多张工作表。
	看下面，然而，只有一张工作表是激活的。
	
	第一张工作表默认是激活的。

  select()

	"select()"方法用于表明从含有多张工作表的工作簿中选取一张：

        $worksheet1->activate();
        $worksheet2->select();
        $worksheet3->select();

	被选中的工作表的标签是高亮的。选取多张工作表是把它们组合在一块的一种方法，所以，例如，可以一举打印多张工作表。通过"activate()"方法被激活的工作表也会被选中。

  hide()
	 "hide()" 方法用于隐藏一个工作表：

        $worksheet2->hide();

  	为了避免使用中间数据或计算结果迷惑用户，你可能想要隐藏一个工作表


	一个隐藏的工作表不能被激活或被选中，所以该方法与"activate()" 和 "select()"是互相排斥的。此外，因为第一张工作表默认是被选中的，你不能不激活另外的工作表而隐藏第一张工作表：

        $worksheet2->activate();
        $worksheet1->hide();

  set_first_sheet()
 	"activate()"方法决定首先选择哪一张工作表。然而，如果有很多张工作表，被选中的工作表可能不会出现在屏幕上。你可以使用"set_first_sheet()"方法选择最左端可见的工作表来避免：

        for ( 1 .. 20 ) {
            $workbook->add_worksheet;
        }

        $worksheet21 = $workbook->add_worksheet();
        $worksheet22 = $workbook->add_worksheet();

        $worksheet21->set_first_sheet();
        $worksheet22->activate();


	该方法不是经常需要。默认值是第一张工作表。

  protect( $password, \%options )
	"protect()" 方法用于防止工作表被修改：

        $worksheet->protect();


	"protect()"方法也会对开启单元格的"locked"和"hidden"属性有影响，如果设置了单元格的"locked"和"hidden"属性的话。一个*locked*的单元格不能够被编辑，并且该属性默认对所有单元格是开启的。一个隐藏的单元格会显示公式的结果但不是公式本身。

    
	查看"protection.pl" 程序和"CELL FORMATTING" 方法的"set_locked" 和 "set_hidden"格式方法。

	你可以选择性地添加一个密码到工作表中：

        $worksheet->protect( 'drowssap' );

	传递一个空字符串''和开启没有密码的保护一样。

 	注意：Excel中工作表级别的密码提供的保护很脆弱。它没有加密你的数据并且很容易被撤销。"Excel::Writer::XLSX"不支持完全的工作表加密，因为它需要一种完全不同的文件格式并且需要花费几个人数月的时间才能实现。

    
	你可以传递一个带有任何一个或全部如下所示键值的散列引用来指定你想保护哪个工作表元素：

        # Default shown.
        %options = (
            objects               => 0,
            scenarios             => 0,
            format_cells          => 0,
            format_columns        => 0,
            format_rows           => 0,
            insert_columns        => 0,
            insert_rows           => 0,
            insert_hyperlinks     => 0,
            delete_columns        => 0,
            delete_rows           => 0,
            select_locked_cells   => 1,
            sort                  => 0,
            autofilter            => 0,
            pivot_tables          => 0,
            select_unlocked_cells => 1,
        );


	上面显示的是默认的布尔值。单个元素的保护可以使用下面的方法：

        $worksheet->protect( 'drowssap', { insert_rows => 1 } );

  set_selection( $first_row, $first_col, $last_row, $last_col )
    This method can be used to specify which cell or cells are selected in a
    worksheet. The most common requirement is to select a single cell, in
    which case $last_row and $last_col can be omitted. The active cell
    within a selected range is determined by the order in which $first and
    $last are specified. It is also possible to specify a cell or a range
    using A1 notation. 
	该方法用于指定在一张工作表中选择哪个或哪些单元格。最常见的需求是选择一个单元格，此时$last_row 和 $last_col 可以省略。在选区内的激活单元格由指定的$first and $last 的顺序决定。也可以使用A1表示法指定单元格或一个范围。

    Examples:

        $worksheet1->set_selection( 3, 3 );          # 1. Cell D4.
        $worksheet2->set_selection( 3, 3, 6, 6 );    # 2. Cells D4 to G7.
        $worksheet3->set_selection( 6, 6, 3, 3 );    # 3. Cells G7 to D4.
        $worksheet4->set_selection( 'D4' );          # Same as 1.
        $worksheet5->set_selection( 'D4:G7' );       # Same as 2.
        $worksheet6->set_selection( 'G7:D4' );       # Same as 3.

	默认的单元格选区是(0,0),'A1'.

  set_row( $row, $height, $format, $hidden, $level, $collapsed )
 
	该方法用于改变行的默认属性。除$row之外，其它参数都是可选的。

	该方法最普通的用法是更改行高：

        $worksheet->set_row( 0, 20 );    # 第一行的行高改为20

	如果你想设置格式时不更改行高，你可以传递"undef"作为行高参数：

        $worksheet->set_row( 0, undef, $format );

	$format参数可以被应用到行中任何没有格式的单元格上，例如：

        $worksheet->set_row( 0, undef, $format1 );    # Set the format for row 1
        $worksheet->write( 'A1', 'Hello' );           # Defaults to $format1
        $worksheet->write( 'B1', 'Hello', $format2 ); # Keeps $format2

	如果你想用这种方法定义一个行格式，你应该在任何调用"write()"方法的行为之前调用该方法。调用该方法以后会覆盖以前指定的任何格式。

   
	如果想隐藏行，$hidden参数应该设置为1。这被用于，例如，在复杂计算中隐藏中间步骤：

        $worksheet->set_row( 0, 20,    $format, 1 );
        $worksheet->set_row( 1, undef, undef,   1 );


	 $level参数用于设置行的分级显示（outline level）。"OUTLINES AND GROUPING IN EXCEL"章节有关于分级显示的描述。有同样分级显示的行会被组合到一块成为一个单一的分级显示。


	下面的例子为行1和行2（从0开始索引）设置了分级显示1：

        $worksheet->set_row( 1, undef, undef, 0, 1 );
        $worksheet->set_row( 2, undef, undef, 0, 1 );

	当和$level参数一同使用的时候，$hidden参数也能用于隐藏折叠的分级显示行：

        $worksheet->set_row( 1, undef, undef, 1, 1 );
        $worksheet->set_row( 2, undef, undef, 1, 1 );


	对于折叠的分级显示，你应该使用可选项$collapsed参数指明哪一行含有折叠符号"+".

        $worksheet->set_row( 3, undef, undef, 0, 0, 1 );

    查看 "outline.pl" 和 "outline_collapsed.pl" 。


	Excel允许多大7个分级显示。因此，$level参数应该在范围："0 <= $level <= 7"内。

  set_column( $first_col, $last_col, $width, $format, $hidden, $level, $collapsed )
  
	该方法用于改变单一列一定范围的列的默认属性。除了 $first_col 和 $last_col 外，所有参数都是可选的。

 	如果"set_column()"被应用到单一列， $first_col和 $last_col的值应该一样。
	在$last_col是0的情况下，它被设置为与$first_col一样的值。

	也可更清晰的使用列的A1表示法来指定列的范围。

    例子:

        $worksheet->set_column( 0, 0, 20 );    # Column  A   width set to 20
        $worksheet->set_column( 1, 3, 30 );    # Columns B-D width set to 30
        $worksheet->set_column( 'E:E', 20 );   # Column  E   width set to 20
        $worksheet->set_column( 'F:H', 30 );   # Columns F-H width set to 30

    The width corresponds to the column width value that is specified in
    Excel. It is approximately equal to the length of a string in the
    default font of Arial 10. Unfortunately, there is no way to specify
    "AutoFit" for a column in the Excel file format. This feature is only
    available at runtime from within Excel.
	

    通常 $format参数是可选的,更多信息,查看"CELL FORMATTING". 
	如果你想设置格式时不更改列宽，你可以传递"undef"作为列宽参数：

        $worksheet->set_column( 0, 0, undef, $format );

	$format参数可以被应用到列中任何没有格式的单元格上，例如：

        $worksheet->set_column( 'A:A', undef, $format1 );    # Set format for col 1
        $worksheet->write( 'A1', 'Hello' );                  # Defaults to $format1
        $worksheet->write( 'A2', 'Hello', $format2 );        # Keeps $format2

	如果你想用这种方法定义一个列格式，你应该在任何调用"write()"方法的行为之前调用该方法。如果你在调用"write()"方法之后调用该方法，它不会有任何效果。

	默认的行格式优于默认的列格式。

        $worksheet->set_row( 0, undef, $format1 );           # Set format for row 1
        $worksheet->set_column( 'A:A', undef, $format2 );    # Set format for col 1
        $worksheet->write( 'A1', 'Hello' );                  # Defaults to $format1
        $worksheet->write( 'A2', 'Hello' );                  # Defaults to $format2

 
	如果想隐藏列，$hidden参数应该设置为1。这被用于，例如，在复杂计算中隐藏中间步骤：


        $worksheet->set_column( 'D:D', 20,    $format, 1 );
        $worksheet->set_column( 'E:E', undef, undef,   1 );

    
	$level参数用于设置列的分级显示（outline level）。"OUTLINES AND GROUPING IN EXCEL"章节有关于分级显示的描述。有同样分级显示的列会被组合到一块成为一个单一的分级显示。

	下面的例子为从B到G的列设置了分级显示1：

        $worksheet->set_column( 'B:G', undef, undef, 0, 1 );


	当和$level参数一同使用的时候，$hidden参数也能用于隐藏折叠的分级显示列：

        $worksheet->set_column( 'B:G', undef, undef, 1, 1 );

  	对于折叠的分级显示，你应该使用可选项$collapsed参数指明哪一行含有折叠符号"+".

        $worksheet->set_column( 'H:H', undef, undef, 0, 0, 1 );

    查看那outline.pl" 和 "outline_collapsed.pl" 程序获取更详细的描述。

	Excel 允许多达7级的分级显示。因此，$level参数应该在范围"0 <= $level <= 7"内。

  outline_settings( $visible, $symbols_below, $symbols_right, $auto_style )

	 "outline_settings()"方法用于控制Excel中分级显示的出现。分级显示在"OUTLINES AND GROUPING IN
    EXCEL"中有描述。

	$visible参数用于控制分级显示是否可见。将该参数设置为0会导致工作表中所有的分级显示被隐藏。你可以使用"Show Outline Symbols"命令按钮将它们显示出来。默认设置为1，即显示分级。

        $worksheet->outline_settings( 0 );

    The $symbols_below parameter is used to control whether the row outline
    symbol will appear above or below the outline level bar. The default
    setting is 1 for symbols to appear below the outline level bar.
	$symbols_below参数用于控制行分级显示标志符是否会出现在分级显示工具条的上方或下面。
	默认的设置为1，即标识符出现在分级显示工具条的下面。

 
	"symbols_right"参数用于控制列分级显示标识符是否会出现在分级显示工具条的左侧或右侧。
	默认设置为1，即标识符出现在分级显示的右边。

    The $auto_style parameter is used to control whether the automatic
    outline generator in Excel uses automatic styles when creating an
    outline. This has no effect on a file generated by "Excel::Writer::XLSX"
    but it does have an effect on how the worksheet behaves after it is
    created. The default setting is 0 for "Automatic Styles" to be turned
    off.
	$auto_style 参数用于控制在Excel中的自动分级显示生成器是否使用自动风格创建分级显示。
	由"Excel::Writer::XLSX"生成的文件间没有区别，但是对于创建后工作表如何表现是有区别的。
	默认设置为0，即关闭"Automatic Styles" 
    The default settings for all of these parameters correspond to Excel's
    default parameters.
    所有这种参数的默认设置与Excel的默认参数有关。

	由 "outline_settings()"方法控制的工作表参数很少使用。

  freeze_panes( $row, $col, $top_row, $left_col )  #冻结窗格
   
	该方法用于将工作表划分为水平或垂直的叫做窗格的区域并且冻结这些窗格以使分隔条不可见。这与Excel中的"窗口->冻结窗格"菜单命令的作用相同。

    The parameters $row and $col are used to specify the location of the
    split. It should be noted that the split is specified at the top or left
    of a cell and that the method uses zero based indexing. Therefore to
    freeze the first row of a worksheet it is necessary to specify the split
    at row 2 (which is 1 as the zero-based index). This might lead you to
    think that you are using a 1 based index but this is not the case.
	$row 和 $col参数用于指定分隔的位置。应该注意的是分隔由单元格的顶部或左边指定，并且该方法使用基于0开始的索引。因此，冻结工作表的第一行，指定第二行（作为基于0的索引时是1）进行分隔是有必要的。这可能导致你认为在使用基于1的索引，但是这不是问题。

    You can set one of the $row and $col parameters as zero if you do not
    want either a vertical or horizontal split.
	如果你不需要水平或垂直分隔，你可以将$row和$col参数中的一个设置为0。

    例子:

        $worksheet->freeze_panes( 1, 0 );    # Freeze the first row
        $worksheet->freeze_panes( 'A2' );    # Same using A1 notation
        $worksheet->freeze_panes( 0, 1 );    # Freeze the first column
        $worksheet->freeze_panes( 'B1' );    # Same using A1 notation
        $worksheet->freeze_panes( 1, 2 );    # Freeze first row and first 2 columns
        $worksheet->freeze_panes( 'C2' );    # Same using A1 notation

    The parameters $top_row and $left_col are optional. They are used to
    specify the top-most or left-most visible row or column in the scrolling
    region of the panes. 例如， to freeze the first row and to have the
    scrolling region begin at row twenty:
	$top_row 和 $left_col参数是可选的。它们用于在窗格的滚动区域中指定可见的最顶端或最左端的行或列。
	例如，冻结第一行并让滚动区域从第20行开始：

        $worksheet->freeze_panes( 1, 0, 20, 0 );

	对于$top_row 和 $left_col参数，你可以使用A1表示法。

    查看那"panes.pl"程序。

  split_panes( $y, $x, $top_row, $left_col )
    This method can be used to divide a worksheet into horizontal or
    vertical regions known as panes. This method is different from the
    "freeze_panes()" method in that the splits between the panes will be
    visible to the user and each pane will have its own scroll bars.
	该方法用于将工作表划分为叫做窗格的水平的或垂直的区域。该方法不同于"freeze_panes()"方法，
	它分隔的窗格对用户是可见的，并且每个窗格都有它们自己的滚动条。

    The parameters $y and $x are used to specify the vertical and horizontal
    position of the split. The units for $y and $x are the same as those
    used by Excel to specify row height and column width. However, the
    vertical and horizontal units are different from each other. Therefore
    you must specify the $y and $x parameters in terms of the row heights
    and column widths that you have set or the default values which are 15
    for a row and 8.43 for a column.
	$y 和 $x 参数用于分隔的水平和垂直位置。 $y和$x的单位与Excel使用的单位一样，用于指定行高和列宽。然而，水平的和垂直的单位不一样。 
	因此，你必须按照你设置好的行高和列宽来指定$y 和 $x 参数，或者，使用默认值，行高15,列宽8.43.

    You can set one of the $y and $x parameters as zero if you do not want
    either a vertical or horizontal split. The parameters $top_row and
    $left_col are optional. They are used to specify the top-most or
    left-most visible row or column in the bottom-right pane.
	如果你不想水平分隔或垂直分隔，你可以将$y 和 $x参数之一设置为0。top_row 和$left_col参数是可选的，它们用于指定在右底部窗格中的最上或最左的可见行或列。

    例子:

        $worksheet->split_panes( 15, 0,   );    # First row
        $worksheet->split_panes( 0,  8.43 );    # First column
        $worksheet->split_panes( 15, 8.43 );    # First row and column

    该方法可以使用A1表示法。

    查看 "freeze_panes()"方法和 "panes.pl"程序。

  merge_range( $first_row, $first_col, $last_row, $last_col, $token, $format )
    The "merge_range()" method allows you merge cells that contain other
    types of alignment in addition to the merging:
	"merge_range()"方法允许你合并含有其它对齐方式（除了合并）的单元格：

        my $format = $workbook->add_format(
            border => 6,
            valign => 'vcenter',
            align  => 'center',
        );

        $worksheet->merge_range( 'B3:D4', 'Vertical and horizontal', $format );

	"merge_range()"方法使用工作表的"write()"方法写入它的$token参数。因此，它会按要求处理数字，字符串，公式或urls。如果你想指定要求的"write_*"方法，使用 "merge_range_type()"方法，看下面。

	查看"merge3.pl"到"merge6.pl"获取该方法完全的信息。

  merge_range_type( $type, $first_row, $first_col, $last_row, $last_col, ... )
    The "merge_range()" method, see above, uses "write()" to insert the
    required data into to a merged range. However, there may be times where
    this isn't what you require so as an alternative the "merge_range_type
    ()" method allows you to specify the type of data you wish to write. For
    example:
	"merge_range()"方法，看上面，使用"write()"向合并的区域中插入需要的数据。然而，有时这可能不是你想要的，所以作为选择，"merge_range_type()"方法允许你指定你想写入的数据类型。比如：

        $worksheet->merge_range_type( 'number',  'B2:C2', 123,    $format1 );
        $worksheet->merge_range_type( 'string',  'B4:C4', 'foo',  $format2 );
        $worksheet->merge_range_type( 'formula', 'B6:C6', '=1+2', $format3 );

    The $type must be one of the following, which corresponds to a
    "write_*()" method:
	$type必须是下面的之一，与 "write_*()"方法相对：

        'number'
        'string'
        'formula'
        'array_formula'
        'blank'
        'rich_string'
        'date_time'
        'url'

    Any arguments after the range should be whatever the appropriate method
    accepts:
	任何在这范围之后的参数应该是任何合适的方法可接受的：

        $worksheet->merge_range_type( 'rich_string', 'B8:C8',
                                      'This is ', $bold, 'bold', $format4 );

    Note, you must always pass a $format object as an argument, even if it
    is a default format.
    注意，你必须一直传递一个$format对象作为参数，即使它是一个默认的格式。
  set_zoom( $scale )
    Set the worksheet zoom factor in the range "10 <= $scale <= 400":
	在范围"10 <= $scale <= 400"内设置工作表的缩放因数：

        $worksheet1->set_zoom( 50 );
        $worksheet2->set_zoom( 75 );
        $worksheet3->set_zoom( 300 );
        $worksheet4->set_zoom( 400 );

    The default zoom factor is 100. You cannot zoom to "Selection" because
    it is calculated by Excel at run-time.
	默认的缩放因数是100.你不能对选取进行缩放，因为它在运行时被Excel计算。

    Note, "set_zoom()" does not affect the scale of the printed page. For
    that you should use "set_print_scale()".
	注意，"set_zoom()" 不影响打印页的尺寸。对于此，你应该使用"set_print_scale()".

  right_to_left()
    The "right_to_left()" method is used to change the default direction of
    the worksheet from left-to-right, with the A1 cell in the top left, to
    right-to-left, with the he A1 cell in the top right.
	"right_to_left()"方法用于改变工作表的默认方向，由从左到右，即A1单元格在左上方，改为从右到左，即A1单元格在右上方。

        $worksheet->right_to_left();

    This is useful when creating Arabic, Hebrew or other near or far eastern
    worksheets that use right-to-left as the default direction.
	当创建阿拉伯的、希伯来的或其它接近东方或远东的默认使用从右到左方向的工作表时有用。
	

  hide_zero()
   
	"hide_zero()"方法用于隐藏任何出现在单元格中的0值。

        $worksheet->hide_zero();

	在Excel中，该选项可在工具->选项->查看菜单下找到。

  set_tab_color()
    The "set_tab_color()" method is used to change the colour of the
    worksheet tab. This feature is only available in Excel 2002 and later.
    You can use one of the standard colour names provided by the Format
    object or a colour index. See "COLOURS IN EXCEL" and the
    "set_custom_color()" method.
	"set_tab_color()"方法用于改变工作表栏的颜色。该功能只在Excel 2002及以后可用。你可以使用格式对象或颜色索引提供的标准颜色名之一。查看"COLOURS IN EXCEL" 和"set_custom_color()"方法。

        $worksheet1->set_tab_color( 'red' );
        $worksheet2->set_tab_color( 0x0C );

    查看"tab_colors.pl" 程序。

  autofilter( $first_row, $first_col, $last_row, $last_col )
    This method allows an autofilter to be added to a worksheet. An
    autofilter is a way of adding drop down lists to the headers of a 2D
    range of worksheet data. This is turn allow users to filter the data
    based on simple criteria so that some data is shown and some is hidden.
	该方法允许向工作表中添加一个自动筛选功能。

    To add an autofilter to a worksheet:
	向工作表添加一个自动筛选：

        $worksheet->autofilter( 0, 0, 10, 3 );
        $worksheet->autofilter( 'A1:D11' );    # Same as above in A1 notation.

    Filter conditions can be applied using the "filter_column()" or
    "filter_column_list()" method.
	筛选条件能使用"filter_column()"或"filter_column_list()"方法应用。

    查看"autofilter.pl" 程序。

  filter_column( $column, $expression )
    The "filter_column" method can be used to filter columns in a autofilter
    range based on simple conditions.
	"filter_column"方法能用于根据简单的条件在一个自动筛选范围内过滤列。

    NOTE: It isn't sufficient to just specify the filter condition. You must
    also hide any rows that don't match the filter condition. Rows are
    hidden using the "set_row()" "visible" parameter. "Excel::Writer::XLSX"
    cannot do this automatically since it isn't part of the file format. See
    the "autofilter.pl" program in the examples directory of the distro for
    an example.
	注意：仅仅指定过滤条件是不够的。你也必须隐藏任何不匹配过滤条件的行。
	使用"set_row()" "visible" 参数来隐藏行。
	"Excel::Writer::XLSX"不能自动地做到这点因为它不是文件格式的一部分。查看"autofilter.pl" 程序。

    The conditions for the filter are specified using simple expressions:
	使用简单的表达式指定过滤条件：

        $worksheet->filter_column( 'A', 'x > 2000' );
        $worksheet->filter_column( 'B', 'x > 2000 and x < 5000' );

    The $column parameter can either be a zero indexed column number or a
    string column name.
	$column参数可以是从0索引的一个列编号或一个字符串列名。

	下面的操作符是可用的：

        操作符        同义词
           ==           =   eq  =~
           !=           <>  ne  !=
           >
           <
           >=
           <=

           and          &&
           or           ||

	操作符的同义词仅仅是让你更舒服地使用表达式的语法糖。有一点很重要：表达式会被Excel解释而不是perl

	一个表达式能由一个单一的语句组成或由"and" 和 "or"操作符分开的2个语句组成，例如：

        'x <  2000'
        'x >  2000'
        'x == 2000'
        'x >  2000 and x <  5000'
        'x == 2000 or  x == 5000'

 
	在表达式中使用"Blanks"或 "NonBlanks"值能达到过滤空开或非空白数据的作用：

        'x == Blanks'
        'x == NonBlanks'

	Excel也允许一些简单的字符串匹配操作：

        'x =~ b*'   # begins with b
        'x !~ b*'   # doesn't begin with b
        'x =~ *b'   # ends with b
        'x !~ *b'   # doesn't end with b
        'x =~ *b*'  # contains b
        'x !~ *b*'  # doesn't contains b


	你可以使用"*"匹配任意字符或数字，使用"?"匹配任一字符或数字。Excel的过滤器不支持其它的正则表达式量词。Excel的正则表达式字符能使用"~"符号转义。

	上面的占位符变量"x"能被任意的简单字符串代替。实际的占位符名在内部被忽略，所以下面的是等价的：

        'x     < 2000'
        'col   < 2000'
        'Price < 2000'

    Also, note that a filter condition can only be applied to a column in a
    range specified by the "autofilter()" Worksheet method.
	注意，过滤条件仅能应用到由"autofilter()"工作表方法所指定范围的列上。

    查看"autofilter.pl" 程序。


	注意Spreadsheet::WriteExcel支持最多10种类型的过滤。这些目前不被 Excel::Writer::XLSX支持，但会在以后添加。

  filter_column_list( $column, @matches )

	在Excel 2007以前只有1到2个过滤条件，比如上面展示的"filter_column" 方法。

    Excel 2007 introduced a new list style filter where it is possible to
    specify 1 or more 'or' style criteria. 例如， if your column
    contained data for the first six months  Then if you selected
    'March', 'April' and 'May' they would be displayed as shown on the
    right.
	Excel 2007引进一种新的列表类型过滤，能指定1个或多个'or'类型的标准。例如，如果你的列中含有前六个月的数据，初始数据会按所有选择the initial data would be displayed as all selected as shown on the left.而如果你选择了 'March', 'April' 和 'May' 它们会显示在右边。


        No criteria selected      Some criteria selected.

        [/] (Select all)          [X] (Select all)
        [/] January               [ ] January
        [/] February              [ ] February
        [/] March                 [/] March
        [/] April                 [/] April
        [/] May                   [/] May
        [/] June                  [ ] June

    The "filter_column_list()" method can be used to represent these types
    of filters:
	"filter_column_list()" 方法能用于代表这些类型的过滤：

        $worksheet->filter_column_list( 'A', 'March', 'April', 'May' );

    The $column parameter can either be a zero indexed column number or a
    string column name.
	$column 参数可以是从0索引的列编号或一个字符串列名。

    可以选择一个或多个标准:

        $worksheet->filter_column_list( 0, 'March' );
        $worksheet->filter_column_list( 1, 100, 110, 120, 130 );

    NOTE: It isn't sufficient to just specify the filter condition. You must
    also hide any rows that don't match the filter condition. Rows are
    hidden using the "set_row()" "visible" parameter. "Excel::Writer::XLSX"
    cannot do this automatically since it isn't part of the file format. See
    the "autofilter.pl" program in the examples directory of the distro for
    an example. e conditions for the filter are specified using simple
    expressions:
	注意：仅仅指定过滤条件是不够的。你也必须隐藏任何不匹配过滤条件的行。
	使用"set_row()" "visible" 参数来隐藏行。
	"Excel::Writer::XLSX"不能自动地做到这点因为它不是文件格式的一部分。查看"autofilter.pl" 程序。

  convert_date_time( $date_string )
  
	"convert_date_time()" 方法在内部被"write_date_time()" 方法使用，用于将日期字符串转换为在Excel中代表日期和时间的数字。

    为了实用目的，它作为一种公共方法显露在我们面前。
	$date_string格式在"write_date_time()"方法中有详细说明。

PAGE SET-UP METHODS
   打印的时侯，页面set-up方法影响一张工作表的外形。它们控制着诸如页眉、页脚和页边距功能。这些方法就是标准的工作表方法。为清晰起见，下面用单独的章节来说明它们的使用。
	下面的方法对于页面设置是可用的：

        set_landscape()
        set_portrait()
        set_page_view()
        set_paper()
        center_horizontally()
        center_vertically()
        set_margins()
        set_header()
        set_footer()
        repeat_rows()
        repeat_columns()
        hide_gridlines()
        print_row_col_headers()
        print_area()
        print_across()
        fit_to_pages()
        set_start_page()
        set_print_scale()
        set_h_pagebreaks()
        set_v_pagebreaks()

    	当使用Excel::Writer::XLSX工作时，通常的需求是将同一个页面设置特性应用到工作簿中的所有工作表中。你可以使用"workbook"类的"sheets()"方法通过访问工作簿中的工作表数组来完成：

        for $worksheet ( $workbook->sheets() ) {
            $worksheet->set_landscape();
        }

  set_landscape()
    This method is used to set the orientation of a worksheet's printed page
    to landscape:
	

        $worksheet->set_landscape();    # Landscape mode

  set_portrait() #设置竖排格式（【印刷】(书页、插图、表格等)竖排格式）
    This method is used to set the orientation of a worksheet's printed page
    to portrait. The default worksheet orientation is portrait, so you won't
    generally need to call this method.
	该方法用于设置工作表打印页面对于竖排的方向。默认的工作表方向是竖排，所以通常你不需要调用该方法。

        $worksheet->set_portrait();    # Portrait mode

  set_page_view()
	该方法用于以"页面查看/布局"模式显示工作表。

        $worksheet->set_page_view();

  set_paper( $index )
    This method is used to set the paper format for the printed output of a
    worksheet. The following paper styles are available:
	该方法用于为工作表的打印输出设置页面格式。下面是可用的纸张类型：

        Index   Paper format            Paper size
        =====   ============            ==========
          0     Printer default         -
          1     Letter                  8 1/2 x 11 in
          2     Letter Small            8 1/2 x 11 in
          3     Tabloid                 11 x 17 in
          4     Ledger                  17 x 11 in
          5     Legal                   8 1/2 x 14 in
          6     Statement               5 1/2 x 8 1/2 in
          7     Executive               7 1/4 x 10 1/2 in
          8     A3                      297 x 420 mm
          9     A4                      210 x 297 mm
         10     A4 Small                210 x 297 mm
         11     A5                      148 x 210 mm
         12     B4                      250 x 354 mm
         13     B5                      182 x 257 mm
         14     Folio                   8 1/2 x 13 in
         15     Quarto                  215 x 275 mm
         16     -                       10x14 in
         17     -                       11x17 in
         18     Note                    8 1/2 x 11 in
         19     Envelope  9             3 7/8 x 8 7/8
         20     Envelope 10             4 1/8 x 9 1/2
         21     Envelope 11             4 1/2 x 10 3/8
         22     Envelope 12             4 3/4 x 11
         23     Envelope 14             5 x 11 1/2
         24     C size sheet            -
         25     D size sheet            -
         26     E size sheet            -
         27     Envelope DL             110 x 220 mm
         28     Envelope C3             324 x 458 mm
         29     Envelope C4             229 x 324 mm
         30     Envelope C5             162 x 229 mm
         31     Envelope C6             114 x 162 mm
         32     Envelope C65            114 x 229 mm
         33     Envelope B4             250 x 353 mm
         34     Envelope B5             176 x 250 mm
         35     Envelope B6             176 x 125 mm
         36     Envelope                110 x 230 mm
         37     Monarch                 3.875 x 7.5 in
         38     Envelope                3 5/8 x 6 1/2 in
         39     Fanfold                 14 7/8 x 11 in
         40     German Std Fanfold      8 1/2 x 12 in
         41     German Legal Fanfold    8 1/2 x 13 in

    Note, it is likely that not all of these paper types will be available
    to the end user since it will depend on the paper formats that the
    user's printer supports. Therefore, it is best to stick to standard
    paper types.
	注意，不是所有的纸张类型对于终端用户都是可用的，因为它依赖于用户的打印机支持的页面格式。因此，最好使用标准的纸张类型。

        $worksheet->set_paper( 1 );    # US Letter
        $worksheet->set_paper( 9 );    # A4

 
	如果你没有指定纸张类型，工作表会使用打印机默认的纸张打印。

  center_horizontally()
    Center the worksheet data horizontally between the margins on the
    printed page:
	在打印页面的页边距之间水平居中对齐工作表数据：

        $worksheet->center_horizontally();

  center_vertically()
    Center the worksheet data vertically between the margins on the printed
    page:
	在打印页的页边距之间垂直居中对齐工作表数据：

        $worksheet->center_vertically();

  set_margins( $inches )
    There are several methods available for setting the worksheet margins on
    the printed page:
	有几种可用的方法用于设置打印页面的工作表页边距：

        set_margins()        # 将所有页边距设为同样的值
        set_margins_LR()     # 将左页边距和右页边距设为同样的值
        set_margins_TB()     # 将上页边距和下页边距设为同样的值
        set_margin_left();   # 设置左页边距
        set_margin_right();  # 设置右页边距
        set_margin_top();    # Set top margin设置上页边距
        set_margin_bottom(); # Set bottom margin设置下页边距

    All of these methods take a distance in inches as a parameter. Note: 1
    inch = 25.4mm. ";-)" The default left and right margin is 0.7 inch. The
    default top and bottom margin is 0.75 inch. Note, these defaults are
    different from the defaults used in the binary file format by
    Spreadsheet::WriteExcel.
	所有这些参数以英寸距离作为参数。注意：1英寸=25.4毫米。默认的左页边距和右页边距是0.7英寸。
	默认的上页边距和下页边距是0.75英寸。注意，这些默认值与Spreadsheet::WriteExcel中使用的二进制文件格式的默认值不同。

  set_header( $string, $margin )
   
	页眉和页脚使用$string生成，$string由普通文本和控制字符组成。 $margin参数是可选的。

	可用的控制字符是：

        控制字符            类别                描述
        =======             ========            ===========
        &L                  对齐                左对齐
        &C                                      居中对齐
        &R                                      右对齐

        &P                  信息                页码
        &N                                      总页数
        &D                                      日期
        &T                                      时间
        &F                                      文件名
        &A                                      工作表名
        &Z                                      工作簿路径

        &fontsize           字体                字体大小
        &"font,style"                           字体名和字体类型
        &U                                      单下划线
        &E                                      双下划线
        &S                                      删除线
        &X                                      上标
        &Y                                      下标

        &&                  其它                字面符号&

	通过在文本前面前置控制字符&L、&C、和&R,可以将页眉和页脚中的文本调整（对齐）为居左、居中和右对齐。

    例如， (使用 ASCII 插图表示结果):

        $worksheet->set_header('&LHello');

         ---------------------------------------------------------------
        |                                                               |
        | Hello                                                         |
        |                                                               |


        $worksheet->set_header('&CHello');

         ---------------------------------------------------------------
        |                                                               |
        |                          Hello                                |
        |                                                               |


        $worksheet->set_header('&RHello');

         ---------------------------------------------------------------
        |                                                               |
        |                                                         Hello |
        |                                                               |

    For simple text, if you do not specify any justification the text will
    be centred. However, you must prefix the text with &C if you specify a
    font name or any other formatting:
	对于纯文本，如果你没有指定任何对齐方式，文本将居中对齐。然而，如果你指定字体名或其它任何格式，你必须在文本前前置&C符号。

        $worksheet->set_header('Hello');

         ---------------------------------------------------------------
        |                                                               |
        |                          Hello                                |
        |                                                               |

	你可以让每个对齐区域中都有文本：

        $worksheet->set_header('&LCiao&CBello&RCielo');

         ---------------------------------------------------------------
        |                                                               |
        | Ciao                     Bello                          Cielo |
        |                                                               |

	当工作表或工作簿发生变化时，信息控制字符作为Excel会更新的变量。时间和日期使用用户默认的格式：

        $worksheet->set_header('&CPage &P of &N');

         ---------------------------------------------------------------
        |                                                               |
        |                        Page 1 of 6                            |
        |                                                               |


        $worksheet->set_header('&CUpdated at &T');

         ---------------------------------------------------------------
        |                                                               |
        |                    Updated at 12:30 PM                        |
        |                                                               |


	你可以通过在字体前前置控制字符&n,"n"是字体大小，来指定文本区域的字体大小：

        $worksheet1->set_header( '&C&30Hello Big' );
        $worksheet2->set_header( '&C&10Hello Small' );


	你可以在文本前前置控制序列&"font,style"来指定文本区域的字体。"font"是诸如"Courier New" 或 "Times New Roman"的字体名，"style"是标准的Windows字体描述之一："Regular", "Italic", "Bold" 或 "Bold Italic":

        $worksheet1->set_header( '&C&"Courier New,Italic"Hello' );
        $worksheet2->set_header( '&C&"Courier New,Bold Italic"Hello' );
        $worksheet3->set_header( '&C&"Times New Roman,Regular"Hello' );

	将所有这些功能组合起来创建复杂的页眉和页脚是可能的。作为对于建立复杂页眉和页脚的帮助，你可以在Excel中记录一个页面设置的宏，并且查看VBA产生的格式字符串。记住VBA使用2个双引号""代表单个双引号".对于上面的最后一个例子，等价的VBA代码看起来像这样：

        .LeftHeader   = ""
        .CenterHeader = "&""Times New Roman,Regular""Hello"
        .RightHeader  = ""

	你应该使用2个and符号"&&"在页眉或页脚中表示一个字面and符号"&":

        $worksheet1->set_header('&CCuriouser && Curiouser - Attorneys at Law');

    As stated above the margin parameter is optional. As with the other
    margins the value should be in inches. The default header and footer
    margin is 0.3 inch. Note, the default margin is different from the
    default used in the binary file format by Spreadsheet::WriteExcel. The
    header and footer margin size can be set as follows:
     上面的例子中，开始页边距参数是可选的。对于其它页边距其值应该是英寸。默认的页眉和页脚页边距是0.3英寸。注意，这些默认页边距值与Spreadsheet::WriteExcel中使用的二进制文件格式的默认值不同。页眉和页脚的页边距大小可以设置如下：
        $worksheet->set_header( '&CHello', 0.75 );

    The header and footer margins are independent of the top and bottom
    margins.
	页眉和页脚的页边距依赖于顶部和底部的页边距。

  
	注意，页眉和页脚字符串必须少于255个字符。长于255个的字符串将不会被写入并生成一个警告。

 
	"set_header()"方法也能处理UTF-8格式的Unicode字符串。

        $worksheet->set_header( "&C\x{263a}" )

    查看 "headers.pl" 程序。

  set_footer()
  
    "set_footer()" 方法与"set_header()"方法的语法一样，看上面。
  repeat_rows( $first_row, $last_row )
    Set the number of rows to repeat at the top of each printed page.
	在每张打印页的顶部设置需要复制的行数。

    For large Excel documents it is often desirable to have the first row or
    rows of the worksheet print out at the top of each page. This can be
    achieved by using the "repeat_rows()" method. The parameters $first_row
    and $last_row are zero based. The $last_row parameter is optional if you
    only wish to specify one row:
	对于很大的Excel文件，在每页的顶部打印工作表的第一行或前几行通常是值得的。这可以使用 "repeat_rows()" 方法做到。$first_row和$last_row参数是基于0的。 如果你只想指定一行，$last_row参数是可选的：

        $worksheet1->repeat_rows( 0 );    # 复制第一行
        $worksheet2->repeat_rows( 0, 1 ); # 复制前2行

  repeat_columns( $first_col, $last_col )
    Set the columns to repeat at the left hand side of each printed page.
	在每张打印页的左边设置需要复制的列数。

    For large Excel documents it is often desirable to have the first column
    or columns of the worksheet print out at the left hand side of each
    page. This can be achieved by using the "repeat_columns()" method. The
    parameters $first_column and $last_column are zero based. The
    $last_column parameter is optional if you only wish to specify one
    column. You can also specify the columns using A1 column notation, see
    the note about "Cell notation".
	
    对于很大的Excel文件，在每页的左侧打印工作表的第一列或前几列通常是值得的。这可以使用 "repeat_columns()" 方法做到。$first_column和$last_column参数是基于0的。 如果你只想指定一列，$last_column参数是可选的。你可以使用A1列表示法指定列数。

        $worksheet1->repeat_columns( 0 );        # Repeat the first column
        $worksheet2->repeat_columns( 0, 1 );     # Repeat the first two columns
        $worksheet3->repeat_columns( 'A:A' );    # Repeat the first column
        $worksheet4->repeat_columns( 'A:B' );    # Repeat the first two columns

  hide_gridlines( $option )
    This method is used to hide the gridlines on the screen and printed
    page. Gridlines are the lines that divide the cells on a worksheet.
     If you have defined your own cell borders you may wish to
    hide the default gridlines.
	该方法用于隐藏屏幕上的网格线和打印过的页面。网格线是在工作表中分隔单元格的线。
	Screen and printed gridlines are turned on by default in an Excel
    worksheet.
   如果你定义了你自己的单元格边框你可能想隐藏默认的网格线。
        $worksheet->hide_gridlines();

	下面的$option值是有效的：

        0 : 不隐藏网格线
        1 : 只隐藏打印后的网格线Hide printed gridlines only
        2 : 隐藏屏幕和打印后的网格线Hide screen and printed gridlines

    If you don't supply an argument or use "undef" the default option is 1,
    i.e. only the printed gridlines are hidden.
	如果你没有提供参数或使用"undef"，则默认的选项是1，i.e。只有打印后的网格线被隐藏。

  print_row_col_headers()
    Set the option to print the row and column headers on the printed page.
	在打印页面上设置选项以打印行标题和列标题。

    An Excel worksheet looks something like the following：
	一张工作表看起来就像如下这样：

         ------------------------------------------
        |   |   A   |   B   |   C   |   D   |  ...
         ------------------------------------------
        | 1 |       |       |       |       |  ...
        | 2 |       |       |       |       |  ...
        | 3 |       |       |       |       |  ...
        | 4 |       |       |       |       |  ...
        |...|  ...  |  ...  |  ...  |  ...  |  ...

    The headers are the letters and numbers at the top and the left of the
    worksheet. Since these headers serve mainly as a indication of position
    on the worksheet they generally do not appear on the printed page. If
    you wish to have them printed you can use the "print_row_col_headers()"
    method :
	标题就是工作表顶部和左侧的字母和数字。因为这些标题主要在工作表中指明位置，它们一般不会出现在打印页面上。如果你想连标题一块打印，你可以使用"print_row_col_headers()"方法。

        $worksheet->print_row_col_headers();

    Do not confuse these headers with page headers as described in the
    "set_header()" section above.
	不要把这些标题和上面提到的有关页面标题的"set_header()"章节弄混淆。

  print_area( $first_row, $first_col, $last_row, $last_col )
    This method is used to specify the area of the worksheet that will be
    printed. All four parameters must be specified. You can also use A1
    notation, 
	该方法用于指定将被打印的工作表区域。4个参数必须都指定。你也可以使用A1表示法：

        $worksheet1->print_area( 'A1:H20' );    # Cells A1 to H20
        $worksheet2->print_area( 0, 0, 19, 7 ); # The same
        $worksheet2->print_area( 'A:H' );       # Columns A to H if rows have data

  print_across()
    The "print_across" method is used to change the default print direction.
    This is referred to by Excel as the sheet "page order".
	"print_across"方法用于改变默认的打印方向。这在Excel工作表中被称为“页面顺序”。

        $worksheet->print_across();

    The default page order is shown below for a worksheet that extends over
    4 pages. The order is called "down then across":
	下面显示的是拥有超过4页的工作表的默认的页面顺序。该顺序叫做“向下然后交叉”

        [1] [3]
        [2] [4]

    However, by using the "print_across" method the print order will be
    changed to "across then down":
	然而，通过使用"print_across"方法，打印顺序会更改为"交叉向下"：

        [1] [2]
        [3] [4]

  fit_to_pages( $width, $height )
    The "fit_to_pages()" method is used to fit the printed area to a
    specific number of pages both vertically and horizontally. If the
    printed area exceeds the specified number of pages it will be scaled
    down to fit. This guarantees that the printed area will always appear on
    the specified number of pages even if the page size or margins change.
	"fit_to_pages()"方法用于垂直和水平地使打印区域与指定页数相适。如果打印区域超过了指定的页数，它会按比列缩小来适应。这保证了即使页面尺寸或页边距发生变化，打印区域也会一直出现在指定页上。

        $worksheet1->fit_to_pages( 1, 1 );    # Fit to 1x1 pages
        $worksheet2->fit_to_pages( 2, 1 );    # Fit to 2x1 pages
        $worksheet3->fit_to_pages( 1, 2 );    # Fit to 1x2 pages

	打印区域能使用上面描述的"print_area()"方法定义。

    A common requirement is to fit the printed output to *n* pages wide but
    have the height be as long as necessary. To achieve this set the $height
    to zero:
    通常的需求是将打印输出为n页宽，但让高度尽可能地长。可以把 $height设置为0来达到要求：
        $worksheet1->fit_to_pages( 1, 0 );    # 1 page wide and as long as necessary

    Note that although it is valid to use both "fit_to_pages()" and
    "set_print_scale()" on the same worksheet only one of these options can
    be active at a time. The last method call made will set the active
    option.
	注意，尽管在同一张工作表中使用"fit_to_pages()" 和 "set_print_scale()" 是正确的，但一次只能激活其中的一个选项。最后一个方法调用会设置激活选项。

    Note that "fit_to_pages()" will override any manual page breaks that are
    defined in the worksheet.
	 注意"fit_to_pages()"会重写任何手册页

  set_start_page( $start_page )
    The "set_start_page()" method is used to set the number of the starting
    page when the worksheet is printed out. The default value is 1.
	 "set_start_page()"方法用于设置工作表打印时的起始页。默认值是1.

        $worksheet->set_start_page( 2 );

  set_print_scale( $scale )
    Set the scale factor of the printed page. Scale factors in the range "10
    <= $scale <= 400" are valid:
	设置打印页的比例系数。在范围"10 <= $scale <= 400"内的比例系数是有效的：

        $worksheet1->set_print_scale( 50 );
        $worksheet2->set_print_scale( 75 );
        $worksheet3->set_print_scale( 300 );
        $worksheet4->set_print_scale( 400 );

    The default scale factor is 100. Note, "set_print_scale()" does not
    affect the scale of the visible page in Excel. For that you should use
    "set_zoom()".
	默认的比例系数是100.注意，"set_print_scale()"不影响Excel可见页的尺寸。对于此。你应使用"set_zoom()"。

    Note also that although it is valid to use both "fit_to_pages()" and
    "set_print_scale()" on the same worksheet only one of these options can
    be active at a time. The last method call made will set the active
    option.
	也要注意，尽管在同一张工作表中使用"fit_to_pages()" 和 "set_print_scale()" 是正确的，但一次只能激活其中的一个选项。最后一个方法调用会设置激活选项。

  set_h_pagebreaks( @breaks )
    Add horizontal page breaks to a worksheet. A page break causes all the
    data that follows it to be printed on the next page. Horizontal page
    breaks act between rows. To create a page break between rows 20 and 21
    you must specify the break at row 21. However in zero index notation
    this is actually row 20. So you can pretend for a small while that you
    are using 1 index notation:
	添加水平分页符到工作表中。分页符导致所有在它后面的数据在下一页中被打印。水平分页符在行之间起作用。为在第20行和第21行之间创建分页符，你必须在第21行指定分页。然而，在以0开始索引的表示法中，这实际上是第20行。所以你可以假装你在使用1索引表示法。

        $worksheet1->set_h_pagebreaks( 20 );    # Break between row 20 and 21

    The "set_h_pagebreaks()" method will accept a list of page breaks and
    you can call it more than once:
	"set_h_pagebreaks()"方法会接受一列分隔符而且你可以多次调用该方法：

        $worksheet2->set_h_pagebreaks( 20,  40,  60,  80,  100 );    # Add breaks
        $worksheet2->set_h_pagebreaks( 120, 140, 160, 180, 200 );    # Add some more

    Note: If you specify the "fit to page" option via the "fit_to_pages()"
    method it will override all manual page breaks.
	注意，如果你通过 "fit_to_pages()"方法指定了"fit to page"选项，它会覆盖所有的手动分页符。

    There is a silent limitation of about 1000 horizontal page breaks per
    worksheet in line with an Excel internal limitation.
	与Excel内部限制一样，每张工作表的水平分页符限制为1000个。

  set_v_pagebreaks( @breaks )
    Add vertical page breaks to a worksheet. A page break causes all the
    data that follows it to be printed on the next page. Vertical page
    breaks act between columns. To create a page break between columns 20
    and 21 you must specify the break at column 21. However in zero index
    notation this is actually column 20. So you can pretend for a small
    while that you are using 1 index notation:
	添加垂直分页符到工作表中。分页符导致所有在它后面的数据在下一页中被打印。垂直分页符在列之间起作用。为在第20列和第21列之间创建分页符，你必须在第21列指定分页。然而，在以0开始索引的表示法中，这实际上是第20列。所以你可以假装你在使用1索引表示法。

        $worksheet1->set_v_pagebreaks(20); #在20和21列之间分页

    The "set_v_pagebreaks()" method will accept a list of page breaks and
    you can call it more than once:
	"set_v_pagebreaks()"方法会接受一列分隔符而且你可以多次调用该方法：


        $worksheet2->set_v_pagebreaks( 20,  40,  60,  80,  100 );    # Add breaks
        $worksheet2->set_v_pagebreaks( 120, 140, 160, 180, 200 );    # Add some more

    Note: If you specify the "fit to page" option via the "fit_to_pages()"
    method it will override all manual page breaks.
	注意，如果你通过 "fit_to_pages()"方法指定了"fit to page"选项，它会覆盖所有的手动分页符。

CELL FORMATTING   #单元格格式化
    This section describes the methods and properties that are available for
    formatting cells in Excel. The properties of a cell that can be
    formatted include: fonts, colours, patterns, borders, alignment and
    number formatting.
	此章节描述在Excel中格式化单元格有哪些方法和属性可用。能用于格式化单元格的属性包括：字体、颜色、样式、边框、对齐和数字格式化。

  创建和使用格式对象
    Cell formatting is defined through a Format object. Format objects are
    created by calling the workbook "add_format()" method as follows:
	单元格的格式是通过格式对象定义的。通过调用如下的工作簿"add_format()"方法创建格式对象：

        my $format1 = $workbook->add_format();            # Set properties later
        my $format2 = $workbook->add_format( %props );    # Set at creation

    The format object holds all the formatting properties that can be
    applied to a cell, a row or a column. The process of setting these
    properties is discussed in the next section.
	格式对象存有所有能应用到单元格的格式属性，一行或一列。在下一章节讨论设置这些属性的步骤。

    Once a Format object has been constructed and its properties have been
    set it can be passed as an argument to the worksheet "write" methods as
    follows:
	一旦创建了格式对象并且设置了它们的属性，它可以按如下方法作为参数传递给工作表的"write"方法：

        $worksheet->write( 0, 0, 'One', $format );
        $worksheet->write_string( 1, 0, 'Two', $format );
        $worksheet->write_number( 2, 0, 3, $format );
        $worksheet->write_blank( 3, 0, $format );

    Formats can also be passed to the worksheet "set_row()" and
    "set_column()" methods to define the default property for a row or
    column.
	格式也可以传递给工作表的"set_row()"和"set_column()"方法为行或列定义默认属性。

        $worksheet->set_row( 0, 15, $format );
        $worksheet->set_column( 0, 0, 15, $format );

  Format methods and Format properties
    格式方法和格式属性
    The following table shows the Excel format categories, the formatting
    properties that can be applied and the equivalent object method:
	下面的表中显示了Excel的格式类别，即能使用的格式属性和等价的对象方法：

        Category   Description       Property        Method Name
        --------   -----------       --------        -----------
        Font       Font type         font            set_font()
                   Font size         size            set_size()
                   Font color        color           set_color()
                   Bold              bold            set_bold()
                   Italic            italic          set_italic()
                   Underline         underline       set_underline()
                   Strikeout         font_strikeout  set_font_strikeout()
                   Super/Subscript   font_script     set_font_script()
                   Outline           font_outline    set_font_outline()
                   Shadow            font_shadow     set_font_shadow()

        Number     Numeric format    num_format      set_num_format()

        Protection Lock cells        locked          set_locked()
                   Hide formulas     hidden          set_hidden()

        Alignment  Horizontal align  align           set_align()
                   Vertical align    valign          set_align()
                   Rotation          rotation        set_rotation()
                   Text wrap         text_wrap       set_text_wrap()
                   Justify last      text_justlast   set_text_justlast()
                   Center across     center_across   set_center_across()
                   Indentation       indent          set_indent()
                   Shrink to fit     shrink          set_shrink()

        Pattern    Cell pattern      pattern         set_pattern()
                   Background color  bg_color        set_bg_color()
                   Foreground color  fg_color        set_fg_color()

        Border     Cell border       border          set_border()
                   Bottom border     bottom          set_bottom()
                   Top border        top             set_top()
                   Left border       left            set_left()
                   Right border      right           set_right()
                   Border color      border_color    set_border_color()
                   Bottom color      bottom_color    set_bottom_color()
                   Top color         top_color       set_top_color()
                   Left color        left_color      set_left_color()
                   Right color       right_color     set_right_color()

    There are two ways of setting Format properties: by using the object
    method interface or by setting the property directly. 例如，, a
    typical use of the method interface would be as follows:
	有2中方法设置格式属性：使用对象方法接口或直接设置属性。例如，下面是典型的方法接口用法：

        my $format = $workbook->add_format();
        $format->set_bold();
        $format->set_color( 'red' );

    By comparison the properties can be set directly by passing a hash of
    properties to the Format constructor:
	通过比较，给格式构造函数传递一个属性散列能直接设置属性：

        my $format = $workbook->add_format( bold => 1, color => 'red' );

    or after the Format has been constructed by means of the
    "set_format_properties()" method as follows:
	或在格式创建之后按下面的方法通过"set_format_properties()"方法设置属性：

        my $format = $workbook->add_format();
        $format->set_format_properties( bold => 1, color => 'red' );

    You can also store the properties in one or more named hashes and pass
    them to the required method:
	你也可以将属性存储在一个或多个具名散列中并将需要的方法传递给它们：

        my %font = (
            font  => 'Arial',
            size  => 12,
            color => 'blue',
            bold  => 1,
        );

        my %shading = (
            bg_color => 'green',
            pattern  => 1,
        );


        my $format1 = $workbook->add_format( %font );            # Font only
        my $format2 = $workbook->add_format( %font, %shading );  # Font and shading

    The method mechanism may be better if you prefer
    setting properties via method calls (which the author did when the code
    was first written) otherwise passing properties to the constructor has
    proved to be a little more flexible and self documenting in practice. An
    additional advantage of working with property hashes is that it allows
    you to share formatting between workbook objects as shown in the example
    above.
	如果你喜欢通过方法调用设置属性，方法机制可能更好。否则，传递属性到构造函数会更复杂些并且它的文档更实用。使用属性散列的一个额外好处是它允许你在上面显示的例子中的工作簿对象之间共享格式。
	

    The Perl/Tk style of adding properties is also supported:
	也支持添加Perl/Tk风格的属性：

        my %font = (
            -font  => 'Arial',
            -size  => 12,
            -color => 'blue',
            -bold  => 1,
        );

  Working with formats使用格式
    The default format is Arial 10 with all other properties off.
	默认的格式是Arial 10，其它属性都关闭。

    Each unique format in Excel::Writer::XLSX must have a corresponding
    Format object. It isn't possible to use a Format with a write() method
    and then redefine the Format for use at a later stage. This is because a
    Format is applied to a cell not in its current state but in its final
    state. Consider the following example:
	在Excel::Writer::XLSX 中，每个单独的格式都有一个相应的格式对象。使用带有write() 方法的格式然后在以后使用阶段再重新定义格式是不可行的。
	这是因为被应用到单元格中的格式不是在它们的当前状态，而是它们的最终状态。看看下面的例子：

        my $format = $workbook->add_format();
        $format->set_bold();
        $format->set_color( 'red' );
        $worksheet->write( 'A1', 'Cell A1', $format );
        $format->set_color( 'green' );
        $worksheet->write( 'B1', 'Cell B1', $format );

    Cell A1 is assigned the Format $format which is initially set to the
    colour red. However, the colour is subsequently set to green. When Excel
    displays Cell A1 it will display the final state of the Format which in
    this case will be the colour green.
	单元格A1被指定格式$format,它开始被设置为红色。然而，颜色随后被设置为绿色。当Excel显示单元格A1时，它会显示格式的最终状态，此处是绿色。

    In general a method call without an argument will turn a property on,
    例如，:
	通常，不带参数的方法调用会开启一个属性，例如：

        my $format1 = $workbook->add_format();
        $format1->set_bold();       # Turns bold on
        $format1->set_bold( 1 );    # Also turns bold on
        $format1->set_bold( 0 );    # Turns bold off

FORMAT METHODS 格式方法
    The Format object methods are described in more detail in the following
    sections. In addition, there is a Perl program called "formats.pl" in
    the "examples" directory of the WriteExcel distribution. This program
    creates an Excel workbook called "formats.xlsx" which contains examples
    of almost all the format types.
	下面的章节详细描述了格式对象方法。此外，有一个叫做"formats.pl"的Perl程序。该程序创建了一个名为"formats.xlsx"的Excel工作簿，它包含了几乎所有格式类型的例子。

	下面的格式方法是可用的：

        set_font()
        set_size()
        set_color()
        set_bold()
        set_italic()
        set_underline()
        set_font_strikeout()
        set_font_script()
        set_font_outline()
        set_font_shadow()
        set_num_format()
        set_locked()
        set_hidden()
        set_align()
        set_rotation()
        set_text_wrap()
        set_text_justlast()
        set_center_across()
        set_indent()
        set_shrink()
        set_pattern()
        set_bg_color()
        set_fg_color()
        set_border()
        set_bottom()
        set_top()
        set_left()
        set_right()
        set_border_color()
        set_bottom_color()
        set_top_color()
        set_left_color()
        set_right_color()

    The above methods can also be applied directly as properties. For
    example "$format->set_bold()" is equivalent to
    "$workbook->add_format(bold => 1)".
	上面的方法能直接用于属性。例如，"$format->set_bold()"方法与 "$workbook->add_format(bold => 1)"方法等价。

  set_format_properties( %properties )
    The properties of an existing Format object can be also be set by means
    of "set_format_properties()":
	通过设置"set_format_properties()"也能设置一个已经存在的格式对象的属性：

        my $format = $workbook->add_format();
        $format->set_format_properties( bold => 1, color => 'red' );

    However, this method is here mainly for legacy reasons. It is preferable
    to set the properties in the format constructor:
	然而，由于历史问题该方法主要用在这里。在格式构造函数中设置属性更合适：

        my $format = $workbook->add_format( bold => 1, color => 'red' );

  set_font( $fontname )
        Default state:      Font is Arial
        Default action:     None
        Valid args:         Any valid font name

     指定使用的字体:

        $format->set_font('Times New Roman');

    Excel can only display fonts that are installed on the system that it is
    running on. Therefore it is best to use the fonts that come as standard
    such as 'Arial', 'Times New Roman' and 'Courier New'. See also the Fonts
    worksheet created by formats.pl
	Excel只能显示安装在系统中正使用着的字体。因此，最好使用作为标准的诸如'Arial', 'Times New Roman' 和 'Courier New'.字体。查看由formats.pl创建的工作表字体。

  set_size()
        Default state:      Font size is 10
        Default action:     Set font size to 1
        Valid args:         Integer values from 1 to as big as your screen.

    Set the font size. Excel adjusts the height of a row to accommodate the
    largest font size in the row. You can also explicitly specify the height
    of a row using the set_row() worksheet method.
	设置字体大小。Excel会调整行高以适应行中的最大字体。你也可以显式地使用set_row()工作表方法指定行高。

        my $format = $workbook->add_format();
        $format->set_size( 30 );

  set_color()
        Default state:      Excels的默认颜色，通常是黑色
        Default action:     设置默认颜色
        Valid args:         8..63之间的整数或下面的字符串:
                            'black'
                            'blue'
                            'brown'
                            'cyan'
                            'gray'
                            'green'
                            'lime'
                            'magenta'
                            'navy'
                            'orange'
                            'pink'
                            'purple'
                            'red'
                            'silver'
                            'white'
                            'yellow'

    
	设置字体颜色。"set_color()"方法用法如下：

        my $format = $workbook->add_format();
        $format->set_color( 'red' );
        $worksheet->write( 0, 0, 'wheelbarrow', $format );

   	注意："set_color()" 方法用于单元格中字体的颜色。要设置单元格的颜色，使用"set_bg_color()" 和
    "set_pattern()" 方法.

	其它的例子，请查看formats.pl程序的'Named colors' 和 'Standard colors'工作表。

    

  set_bold()
        Default state:      bold is off
        Default action:     Turn bold on
        Valid args:         0, 1

    
	设置字体的bold黑体属性：

        $format->set_bold();  # Turn bold on

  set_italic()
        Default state:      Italic is off
        Default action:     Turn italic on
        Valid args:         0, 1


	设置字体的斜体属性：

        $format->set_italic();  # Turn italic on

  set_underline()
        Default state:      Underline is off
        Default action:     Turn on single underline
        Valid args:         0  = 没有下划线
                            1  = 单一下划线
                            2  = 双下划线
                            33 = Single accounting underline
                            34 = Double accounting underline

	设置字体的下划线属性。

        $format->set_underline();   # Single underline

  set_font_strikeout()
        Default state:      Strikeout is off
        Default action:     Turn strikeout on
        Valid args:         0, 1

	设置字体的删除线属性。

  set_font_script()
        Default state:      Super/Subscript is off
        Default action:     Turn Superscript on
        Valid args:         0  = Normal
                            1  = Superscript
                            2  = Subscript

   	设置字体的上标/下标属性。

  set_font_outline()
        Default state:      Outline is off
        Default action:     Turn outline on
        Valid args:         0, 1

    仅支持Mac.

  set_font_shadow()
        Default state:      Shadow is off
        Default action:     Turn shadow on
        Valid args:         0, 1

    Mac only.

  set_num_format()
        Default state:      General format
        Default action:     Format index 1
        Valid args:         See the following table

    This method is used to define the numerical format of a number in Excel.
    It controls whether a number is displayed as an integer, a floating
    point number, a date, a currency value or some other user defined
    format.
	该方法用于定义Excel中数字的数字格式。它控制着一个数字是否显示为整数、浮点数、日期、货币值或其它用户定义的格式。

    The numerical format of a cell can be specified by using a format string
    or an index to one of Excel's built-in formats:
	单元格的数字格式能使用一个格式化字符串或Excel的内建格式索引指定：

        my $format1 = $workbook->add_format();
        my $format2 = $workbook->add_format();
        $format1->set_num_format( 'd mmm yyyy' );    # Format string
        $format2->set_num_format( 0x0f );            # Format index

        $worksheet->write( 0, 0, 36892.521, $format1 );    # 1 Jan 2001
        $worksheet->write( 0, 0, 36892.521, $format2 );    # 1-Jan-01

   	使用格式化字符串能定义很复杂的数字格式.

        $format01->set_num_format( '0.000' );
        $worksheet->write( 0, 0, 3.1415926, $format01 );    # 3.142

        $format02->set_num_format( '#,##0' );
        $worksheet->write( 1, 0, 1234.56, $format02 );      # 1,235

        $format03->set_num_format( '#,##0.00' );
        $worksheet->write( 2, 0, 1234.56, $format03 );      # 1,234.56

        $format04->set_num_format( '$0.00' );
        $worksheet->write( 3, 0, 49.99, $format04 );        # $49.99

      	#注意，你也可以使用其它诸如英镑或日元的货币符号。
        #其它货币可能要求使用Unicode。
        $format07->set_num_format( 'mm/dd/yy' );
        $worksheet->write( 6, 0, 36892.521, $format07 );    # 01/01/01

        $format08->set_num_format( 'mmm d yyyy' );
        $worksheet->write( 7, 0, 36892.521, $format08 );    # Jan 1 2001

        $format09->set_num_format( 'd mmmm yyyy' );
        $worksheet->write( 8, 0, 36892.521, $format09 );    # 1 January 2001

        $format10->set_num_format( 'dd/mm/yyyy hh:mm AM/PM' );
        $worksheet->write( 9, 0, 36892.521, $format10 );    # 01/01/2001 12:30 AM

        $format11->set_num_format( '0 "dollar and" .00 "cents"' );
        $worksheet->write( 10, 0, 1.87, $format11 );        # 1 dollar and .87 cents

        # Conditional numerical formatting.
        $format12->set_num_format( '[Green]General;[Red]-General;General' );
        $worksheet->write( 11, 0, 123, $format12 );         # > 0 Green
        $worksheet->write( 12, 0, -45, $format12 );         # < 0 Red
        $worksheet->write( 13, 0, 0,   $format12 );         # = 0 Default colour

        # 邮政编码
        $format13->set_num_format( '00000' );
        $worksheet->write( 14, 0, '01209', $format13 );

  
	颜色的格式应该使用下列值之一：

        [Black] [Blue] [Cyan] [Green] [Magenta] [Red] [White] [Yellow]

    Alternatively you can specify the colour based on a colour index as
    follows: "[Color n]", where n is a standard Excel colour index - 7. See
    the 'Standard colors' worksheet created by formats.pl.
	作为选择，你可以根据下面的颜色索引指定颜色： "[Color n]"，n是标准的Excel颜色索引-7.查看由formats.pl创建的工作表中的'Standard colors'。

   
    You should ensure that the format string is valid in Excel prior to
    using it in WriteExcel.
	你应该确保格式字符串在 WriteExcel中使用它之前是合法的.

    下面的表显式了Excel内建的格式：
        Index   Index   Format String
        0       0x00    General
        1       0x01    0
        2       0x02    0.00
        3       0x03    #,##0
        4       0x04    #,##0.00
        5       0x05    ($#,##0_);($#,##0)
        6       0x06    ($#,##0_);[Red]($#,##0)
        7       0x07    ($#,##0.00_);($#,##0.00)
        8       0x08    ($#,##0.00_);[Red]($#,##0.00)
        9       0x09    0%
        10      0x0a    0.00%
        11      0x0b    0.00E+00
        12      0x0c    # ?/?
        13      0x0d    # ??/??
        14      0x0e    m/d/yy
        15      0x0f    d-mmm-yy
        16      0x10    d-mmm
        17      0x11    mmm-yy
        18      0x12    h:mm AM/PM
        19      0x13    h:mm:ss AM/PM
        20      0x14    h:mm
        21      0x15    h:mm:ss
        22      0x16    m/d/yy h:mm
        ..      ....    ...........
        37      0x25    (#,##0_);(#,##0)
        38      0x26    (#,##0_);[Red](#,##0)
        39      0x27    (#,##0.00_);(#,##0.00)
        40      0x28    (#,##0.00_);[Red](#,##0.00)
        41      0x29    _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
        42      0x2a    _($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)
        43      0x2b    _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
        44      0x2c    _($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
        45      0x2d    mm:ss
        46      0x2e    [h]:mm:ss
        47      0x2f    mm:ss.0
        48      0x30    ##0.0E+0
        49      0x31    @

	查看 formats.pl.中的'Numerical formats'工作表。
	还可查看number_formats1.html 和number_formats2.html 

    Note 1. Numeric formats 23 to 36 are not documented by Microsoft and may
    differ in international versions.
	注意，23到36之间的数字格式在Microsoft中没有说明文档并且内部之间的版本可能不同。

    Note 2. In Excel 5 the dollar sign appears as a dollar sign. In Excel
    97-2000 it appears as the defined local currency symbol.
	注意2，在Excel5中美元符号以美元符号出现。在Excel97-2000中作为当地定义的货币符号出现。

  set_locked()
        Default state:      Cell locking is on
        Default action:     Turn locking on
        Valid args:         0, 1

    This property can be used to prevent modification of a cells contents.
    Following Excel's convention, cell locking is turned on by default.
    However, it only has an effect if the worksheet has been protected, see
    the worksheet "protect()" method.
	该属性用于防止修改单元格的内容。根据Excel的约定，单元格默认被锁定。然而，只有当工作表被保护时，它才起作用。查看"protect()"方法。

        my $locked = $workbook->add_format();
        $locked->set_locked( 1 );    # A non-op

        my $unlocked = $workbook->add_format();
        $locked->set_locked( 0 );

        # 开启工作表保护
        $worksheet->protect();

        # 该单元格不能被编辑.
        $worksheet->write( 'A1', '=1+2', $locked );

        # 这个单元格能被编辑.
        $worksheet->write( 'A2', '=1+2', $unlocked );

  
	注意：即使有密码这也仅提供了很弱的保护，查看与"protect()" 方法有关的注意事项。

  set_hidden()
        Default state:      Formula hiding is off
        Default action:     Turn hiding on
        Valid args:         0, 1

    This property is used to hide a formula while still displaying its
    result. This is generally used to hide complex calculations from end
    users who are only interested in the result. It only has an effect if
    the worksheet has been protected, see the worksheet "protect()" method.
	该属性用于隐藏一个公式但仍然显示该公式的结果。这通常用于对只关心结果的终端用户隐藏复杂的计算过程。只有当工作表开启保护时，该方法才起作用，查看"protect()" 方法。

        my $hidden = $workbook->add_format();
        $hidden->set_hidden();

        # 开启工作表保护
        $worksheet->protect();

        # 在这个单元格中的公式不可见
        $worksheet->write( 'A1', '=1+2', $hidden );

   注意：即使使用密码，这也仅仅提供了很弱的保护，查看关于"protect()"方法的注意事项。

  set_align()
        Default state:      Alignment is off
        Default action:     Left alignment
        Valid args:         'left'              Horizontal
                            'center'
                            'right'
                            'fill'
                            'justify'
                            'center_across'

                            'top'               Vertical
                            'vcenter'
                            'bottom'
                            'vjustify'

    This method is used to set the horizontal and vertical text alignment
    within a cell. Vertical and horizontal alignments can be combined. The
    method is used as follows:
	该方法用于在单元格中设置文本的水平和垂直对齐方式。垂直和水平对齐方式可以结合。该方法用法如下：

        my $format = $workbook->add_format();
        $format->set_align( 'center' );
        $format->set_align( 'vcenter' );
        $worksheet->set_row( 0, 30 );
        $worksheet->write( 0, 0, 'X', $format );

    Text can be aligned across two or more adjacent cells using the
    "center_across" property. However, for genuine merged cells it is better
    to use the "merge_range()" worksheet method.
	使用"center_across"属性，文本可以在2个或更多相邻的单元格之间对齐。然而，对于真实的合并单元格最好使用"merge_range()"工作表方法。

    The "vjustify" (vertical justify) option can be used to provide
    automatic text wrapping in a cell. The height of the cell will be
    adjusted to accommodate the wrapped text. To specify where the text
    wraps use the "set_text_wrap()" method.
	在单元格中，"vjustify"（垂直调整）选项能用于提供自动文本环绕。单元格的高度会被自动调整以适应环绕文本。使用"set_text_wrap()"方法指定文本环绕的位置。

   	查看formats.pl生成的'Alignment' 工作表获取更多实例。

  set_center_across()
        Default state:      Center across selection is off
        Default action:     Turn center across on
        Valid args:         1

    Text can be aligned across two or more adjacent cells using the
    "set_center_across()" method. This is an alias for the
    "set_align('center_across')" method call.
	使用"set_center_across()"方法，文本能够在2个或多个相邻单元格之间对齐。这是 "set_align('center_across')" 方法调用的别名。

    Only one cell should contain the text, the other cells should be blank:
	应该只有一个单元格包含文本，其它单元格是空的：

        my $format = $workbook->add_format();
        $format->set_center_across();

        $worksheet->write( 1, 1, 'Center across selection', $format );
        $worksheet->write_blank( 1, 2, $format );

    See also the "merge1.pl" to "merge6.pl" programs in the "examples"
    directory and the "merge_range()" method.
	查看"merge1.pl" 到"merge6.pl" 程序和"merge_range()"方法。

  set_text_wrap()
        Default state:      Text wrap is off
        Default action:     Turn text wrap on
        Valid args:         0, 1

    这儿有个例子 using the text wrap property, the escape character
    "\n" is used to indicate the end of line:
	这儿有个使用文本环绕属性的例子，

        my $format = $workbook->add_format();
        $format->set_text_wrap();
        $worksheet->write( 0, 0, "It's\na bum\nwrap", $format );

    Excel will adjust the height of the row to accommodate the wrapped text.
    A similar effect can be obtained without newlines using the
    "set_align('vjustify')" method. See the "textwrap.pl" program in the
    "examples" directory.
	Excel会调整行高以适应环绕文本。使用"set_align('vjustify')" 方法不换行就能获得相似的效果。查看"textwrap.pl"程序。

  set_rotation()
        Default state:      Text rotation is off
        Default action:     None
        Valid args:         Integers in the range -90 to 90 and 270

    Set the rotation of the text in a cell. The rotation can be any angle in
    the range -90 to 90 degrees.
	在Excel中设置文本旋转。旋转的度数可以在-90度到90度之间。

        my $format = $workbook->add_format();
        $format->set_rotation( 30 );
        $worksheet->write( 0, 0, 'This text is rotated', $format );

    The angle 270 is also supported. This indicates text where the letters
    run from top to bottom.
	也支持270度旋转。这表明文本中的字母从顶部旋转到底部。？

  set_indent()
        Default state:      Text indentation is off
        Default action:     Indent text 1 level
        Valid args:         Positive integers

    This method can be used to indent text. The argument, which should be an
    integer, is taken as the level of indentation:
	该方法用于缩排文本。参数应该是一个整数，作为缩进的级别：

        my $format = $workbook->add_format();
        $format->set_indent( 2 );
        $worksheet->write( 0, 0, 'This text is indented', $format );

    Indentation is a horizontal alignment property. It will override any
    other horizontal properties but it can be used in conjunction with
    vertical properties.
	缩进是水平对齐属性。它会覆盖其它任何水平属性但它能与垂直属性一起使用。

  set_shrink()
        Default state:      Text shrinking is off
        Default action:     Turn "shrink to fit" on
        Valid args:         1

    This method can be used to shrink text so that it fits in a cell.
	该方法用于收缩文本以适应单元格的大小。

        my $format = $workbook->add_format();
        $format->set_shrink();
        $worksheet->write( 0, 0, 'Honey, I shrunk the text!', $format );

  set_text_justlast()
        Default state:      Justify last is off
        Default action:     Turn justify last on
        Valid args:         0, 1

    Only applies to Far Eastern versions of Excel.
	只对远东版本的Excel适用。

  set_pattern()
        默认状态:      Pattern is off
        默认行为:      Solid fill is on
        合法参数:      0 .. 18

    Set the background pattern of a cell.
	设置单元格的背景图案。

    Examples of the available patterns are shown in the 'Patterns' worksheet
    created by formats.pl. However, it is unlikely that you will ever need
    anything other than Pattern 1 which is a solid fill of the background
    color.
	可用图案的例子显示在formats.pl创建的'Patterns'工作表中。然而，除了图案1是背景色的完全填充外你不在需要其它东西是不可能的。？

  set_bg_color()
        Default state:      Color is off
        Default action:     Solid fill.
        Valid args:         See set_color()

    The "set_bg_color()" method can be used to set the background colour of
    a pattern. Patterns are defined via the "set_pattern()" method. If a
    pattern hasn't been defined then a solid fill pattern is used as the
    default.
	"set_bg_color()"方法用于设置图案的背景颜色。图案通过"set_pattern()"方法定义。如果没有定义图案，则默认使用完全填充图案。

    这儿有个怎样在单元格中设置完全填充的例子：
	of how to set up a solid fill in a cell:

        my $format = $workbook->add_format();

        $format->set_pattern();    # 使用完全填充是这是可选的

        $format->set_bg_color( 'green' );
        $worksheet->write( 'A1', 'Ray', $format );

	查看formats.pl程序中的'Patterns'工作表获取更多例子。

  set_fg_color()
        Default state:      Color is off
        Default action:     Solid fill.
        Valid args:         See set_color()

    "set_fg_color()"方法用于设置图案的前景色。
    查看formats.pl程序中的'Patterns'工作表获取更多例子。

  set_border()
        Also applies to:    set_bottom()
                            set_top()
                            set_left()
                            set_right()

        Default state:      Border is off
        Default action:     Set border type 1
        Valid args:         0-13, See below.

	单元格边框由底部的、顶部的、左侧的、右侧的边框组成。这些边框能使用"set_border()"设置为同样的颜色，或单独使用上面展示的相关方法调用。

  
	下面显示了由Excel::Writer::XLSX按索引号排序后的边框样式：

        Index   Name            Weight   Style
        =====   =============   ======   ===========
        0       None            0
        1       Continuous      1        -----------
        2       Continuous      2        -----------
        3       Dash            1        - - - - - -
        4       Dot             1        . . . . . .
        5       Continuous      3        -----------
        6       Double          3        ===========
        7       Continuous      0        -----------
        8       Dash            2        - - - - - -
        9       Dash Dot        1        - . - . - .
        10      Dash Dot        2        - . - . - .
        11      Dash Dot Dot    1        - . . - . .
        12      Dash Dot Dot    2        - . . - . .
        13      SlantDash Dot   2        / - . / - .

	下面显示了按样式排序后的边框：

        Name            Weight   Style         Index
        =============   ======   ===========   =====
        Continuous      0        -----------   7
        Continuous      1        -----------   1
        Continuous      2        -----------   2
        Continuous      3        -----------   5
        Dash            1        - - - - - -   3
        Dash            2        - - - - - -   8
        Dash Dot        1        - . - . - .   9
        Dash Dot        2        - . - . - .   10
        Dash Dot Dot    1        - . . - . .   11
        Dash Dot Dot    2        - . . - . .   12
        Dot             1        . . . . . .   4
        Double          3        ===========   6
        None            0                      0
        SlantDash Dot   2        / - . / - .   13

	下面显式了在Excel对话框中排序后的边框：

        Index   Style             Index   Style
        =====   =====             =====   =====
        0       None              12      - . . - . .
        7       -----------       13      / - . / - .
        4       . . . . . .       10      - . - . - .
        11      - . . - . .       8       - - - - - -
        9       - . - . - .       2       -----------
        3       - - - - - -       5       -----------
        1       -----------       6       ===========

    Examples of the available border styles are shown in the 'Borders'
    worksheet created by formats.pl.
	可用的边框样式的例子显示在由formats.pl创建的'Borders'工作表中。

  set_border_color()
        Also applies to:    set_bottom_color()
                            set_top_color()
                            set_left_color()
                            set_right_color()

        Default state:      Color is off
        Default action:     Undefined
        Valid args:         See set_color()


	设置单元格边框的颜色。单元格边框由底边框、顶边框、左边框和右边框组成。
	这些边框能使用"set_border()"设置为同样的颜色，或单独使用上面展示的相关方法调用。
	边框样式和颜色的例子显示在由formats.pl程序创建的 'Borders'工作表中。

  copy( $format )
    This method is used to copy all of the properties from one Format object
    to another:
	该方法用于从一个格式对象中复制所有的属性到另一个格式对象中：

        my $lorry1 = $workbook->add_format();
        $lorry1->set_bold();
        $lorry1->set_italic();
        $lorry1->set_color( 'red' );    # lorry1 is bold, italic and red

        my $lorry2 = $workbook->add_format();
        $lorry2->copy( $lorry1 );
        $lorry2->set_color( 'yellow' );    # lorry2 is bold, italic and yellow

    The "copy()" method is only useful if you are using the method interface
    to Format properties. It generally isn't required if you are setting
    Format properties directly using hashes.
	"copy()"方法只有在你使用该方法接口的格式属性是有用的。如果你直接使用散列设置格式的属性，那一般不需要copy()方法。

 
	注意：这不是一个复制构造函数，在复制之前2个对象都必须是存在的。

UNICODE IN EXCEL

	下面是在 "Excel::Writer::XLSX"中处理Unicode的简介。

   
	Excel::Writer::XLSX与Spreadsheet::WriteExcel 的写入方式不同，后者只处理UTF-8格式的Unicode数据，并且不能处理遗留的UTF-16的Excel格式。

	如果数据是UTF-8格式的，则 Excel::Writer::XLSX 会自动处理它。

    如果你处理的是非UTF-8格式的non-ASCII字符，则perl会提供有用的Encode工具模块帮助你转换为需要的格式，例如：

        use Encode 'decode';

        my $string = 'some string with koi8-r characters';
           $string = decode('koi8-r', $string); # koi8-r to utf8

    Alternatively you can read data from an encoded file and convert it to
    "UTF-8" as you read it in:
	作为选择，当你读入数据时，你能从一个编码后文件中读取数据并将数据转换为UTF-8：

        my $file = 'unicode_koi8r.txt';
        open FH, '<:encoding(koi8-r)', $file or die "Couldn't open $file: $!\n";

        my $row = 0;
        while ( <FH> ) {
            # Data read in is now in utf8 format.
            chomp;
            $worksheet->write( $row++, 0, $_ );
        }

    也请查看"unicode_*.pl"程序。

COLOURS IN EXCEL
    Excel provides a colour palette of 56 colours. In Excel::Writer::XLSX
    these colours are accessed via their palette index in the range 8..63.
    This index is used to set the colour of fonts, cell patterns and cell
    borders. 例如，:
	Excel提供了56种颜色的调色板。在Excel::Writer::XLSX中这些颜色通过它们的颜料索值引来访问，索引值范围是8..63。此处的索引值用于设置字体颜色、单元格图案和单元格边框。例如：

        my $format = $workbook->add_format(
                                            color => 12, # index for blue
                                            font  => 'Arial',
                                            size  => 12,
                                            bold  => 1,
                                         );


	最常用的颜色也能通过它们的名字访问。名字作为颜色索引的简单别名：

        black     =>    8
        blue      =>   12
        brown     =>   16
        cyan      =>   15
        gray      =>   23
        green     =>   17
        lime      =>   11
        magenta   =>   14
        navy      =>   18
        orange    =>   53
        pink      =>   33
        purple    =>   20
        red       =>   10
        silver    =>   22
        white     =>    9
        yellow    =>   13

    例如:

        my $font = $workbook->add_format( color => 'red' );

 	Excel的VBA用户应该注意等价的颜色索引是1..56而非8..63.

    If the default palette does not provide a required colour you can
    override one of the built-in values. This is achieved by using the
    "set_custom_color()" workbook method to adjust the RGB (red green blue)
    components of the colour:
	如果默认的颜料不能提供你想要的颜色，你可以重写其中的内置值。使用"set_custom_color()"工作簿方法来调整颜色的RGB（红 绿 蓝）成分可以做到这点：
        my $ferrari = $workbook->set_custom_color( 40, 216, 12, 12 );

        my $format = $workbook->add_format(
            bg_color => $ferrari,
            pattern  => 1,
            border   => 1
        );

        $worksheet->write_blank( 'A1', $format );

    查看"colors.pl" 程序。
DATES AND TIME IN EXCEL
    
	理解Excel中的日期和时间，有2件重要的事情：

    1、Excel中的日期/时间 是一个实数加上一个Excel数字格式。
   	2、Excel::Writer::XLSX 中的"write()"方法不会自动将“日期/时间”字符串转换成Excel的‘日期/时间’。
    伴随下面的关于怎样将时间和日期转换成需要的格式的一些建议，这2点会有更详细的解释。

	Excel的“日期/时间”就是数字加上格式
    If you write a date string with "write()" then all you will get is a
    string:
	如果你使用"write()"方法写入日期字符串，则所有你将得到的会是一个字符串：

        $worksheet->write( 'A1', '02/03/04' );   # !! 写入一个字符串而非一个日期. !!

 	Excel中日期和数字代表实数，例如，"Jan 1 2001 12:30 AM"代表数字36892.521.

	数据的整数部分存储的是自纪元以来的天数，小数部分存储的是一天的百分比。

    Excel中的日期或时间与其它任何数字相似。为了让数字以日期的形式显示，你必须将一个Excel数字格式应用到这个数字上。下面是一些例子：

        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        my $workbook  = Excel::Writer::XLSX->new( 'date_examples.xlsx' );
        my $worksheet = $workbook->add_worksheet();

        $worksheet->set_column( 'A:A', 30 );    # For extra visibility.

        my $number = 39506.5;

        $worksheet->write( 'A1', $number );             #   39506.5

        my $format2 = $workbook->add_format( num_format => 'dd/mm/yy' );
        $worksheet->write( 'A2', $number, $format2 );    #  28/02/08

        my $format3 = $workbook->add_format( num_format => 'mm/dd/yy' );
        $worksheet->write( 'A3', $number, $format3 );    #  02/28/08

        my $format4 = $workbook->add_format( num_format => 'd-m-yyyy' );
        $worksheet->write( 'A4', $number, $format4 );    #  28-2-2008

        my $format5 = $workbook->add_format( num_format => 'dd/mm/yy hh:mm' );
        $worksheet->write( 'A5', $number, $format5 );    #  28/02/08 12:00

        my $format6 = $workbook->add_format( num_format => 'd mmm yyyy' );
        $worksheet->write( 'A6', $number, $format6 );    # 28 Feb 2008

        my $format7 = $workbook->add_format( num_format => 'mmm d yyyy hh:mm AM/PM' );
        $worksheet->write('A7', $number , $format7);     #  Feb 28 2008 12:00 PM

  Excel::Writer::XLSX 不自动转换“日期/时间”字符串
    Excel::Writer::XLSX doesn't automatically convert input date strings
    into Excel's formatted date numbers due to the large number of possible
    date formats and also due to the possibility of misinterpretation.
	由于可用的日期格式数量很大，也由于可能的误解，Excel::Writer::XLSX不能将输入的日期字符串自动转换为Excel的格式化日期数字。

    例如，, does "02/03/04" mean March 2 2004, February 3 2004 or even
    March 4 2002.
	例如，"02/03/04"的意思是 March 2 2004, February 3 2004 甚至是 March 4 2002吗？

    Therefore, in order to handle dates you will have to convert them to
    numbers and apply an Excel format. Some methods for converting dates are
    listed in the next section.
	因此，为了处理日期你必须将它们转换为数字并应用一个Excel格式。转换日期的一些方法在下面的章节中列出。


    最直接的方式是将你的数据转换为ISO8601"yyyy-mm-ddThh:mm:ss.sss" 日期格式，并使用"write_date_time()"工作表方法:
        $worksheet->write_date_time( 'A2', '2001-01-01T12:20', $format );

    查看文档的"write_date_time()"章节获取详细信息。

  
	处理日期字符的一般方法是使用"write_date_time()"：

        1.使用正则表达式识别输入的日期/时间。
		2.使用同样的正则表达式提取日期/时间的组成部分。
		3.将日期/时间转换为ISO8601格式。
        4.使用 write_date_time()和数字格式写入日期/时间。

    这儿有个例子:

        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        my $workbook  = Excel::Writer::XLSX->new( 'example.xlsx' );
        my $worksheet = $workbook->add_worksheet();

        # 为数据设置默认格式
        my $date_format = $workbook->add_format( num_format => 'mmm d yyyy' );

        # Increase column width to improve visibility of data.
        $worksheet->set_column( 'A:C', 20 );

        # Simulate reading from a data source.
        my $row = 0;

        while ( <DATA> ) {
            chomp;

            my $col  = 0;
            my @data = split ' ';

            for my $item ( @data ) {

                # Match dates in the following formats: d/m/yy, d/m/yyyy
                if ( $item =~ qr[^(\d{1,2})/(\d{1,2})/(\d{4})$] ) {

                    # Change to the date format required by write_date_time().
                    my $date = sprintf "%4d-%02d-%02dT", $3, $2, $1;

                    $worksheet->write_date_time( $row, $col++, $date,
                        $date_format );
                }
                else {

                    # Just plain data
                    $worksheet->write( $row, $col++, $item );
                }
            }
            $row++;
        }

        __DATA__
        Item    Cost    Date
        Book    10      1/9/2007
        Beer    4       12/9/2007
        Bed     500     5/10/2007


	更高级的方法你可以通过"add_write_handler()"方法修改"write()"方法来处理你选择的数据格式。查看"add_write_handler()"章节和write_handler3.pl、 write_handler4.pl程序。

  Converting dates and times to an Excel date or time
  将日期和时间转换为Excel的日期或时间

	上面的"write_date_time()" 方法只是处理日期和时间的方法之一。

    You can also use the "convert_date_time()" worksheet method to convert
    from an ISO8601 style date string to an Excel date and time number.
	你也可以使用"convert_date_time()"工作表方法将ISO8601风格的字符串转换为Excel的日期和时间数字。

     Excel::Writer::XLSX::Utility模块有日期时间处理函数：

        use Excel::Writer::XLSX::Utility;

        $date           = xl_date_list(2002, 1, 1);         # 37257
        $date           = xl_parse_date("11 July 1997");    # 35622
        $time           = xl_parse_time('3:21:36 PM');      # 0.64
        $date           = xl_decode_date_EU("13 May 2002"); # 37389

    注意，有些函数需要额外的CPAN模块。

    
OUTLINES AND GROUPING IN EXCEL   Excel中的组和分级显示
    Excel allows you to group rows or columns so that they can be hidden or
    displayed with a single mouse click. This feature is referred to as
    outlines.
	Excel允许你将行或列分组，以使单击鼠标时它们能被隐藏或显示。该功能叫做分级显示。

    Outlines can reduce complex data down to a few salient sub-totals or
    summaries.
	分级显示能将复杂数据减少到几个突出的小计或总结。

    This feature is best viewed in Excel but the following is an ASCII
    representation of what a worksheet with three outlines might look like.
    Rows 3-4 and rows 7-8 are grouped at level 2. Rows 2-9 are grouped at
    level 1. The lines at the left hand side are called outline level bars.
	该功能最好在Excel中查看，但下面的是带有3个分级显示的工作表所代表的ASCII图。
	3-4行与7-8行在第2级别分组。2-9行在第一级别被分组。左侧的线叫做分级显示条。

                ------------------------------------------
         1 2 3 |   |   A   |   B   |   C   |   D   |  ...
                ------------------------------------------
          _    | 1 |   A   |       |       |       |  ...
         |  _  | 2 |   B   |       |       |       |  ...
         | |   | 3 |  (C)  |       |       |       |  ...
         | |   | 4 |  (D)  |       |       |       |  ...
         | -   | 5 |   E   |       |       |       |  ...
         |  _  | 6 |   F   |       |       |       |  ...
         | |   | 7 |  (G)  |       |       |       |  ...
         | |   | 8 |  (H)  |       |       |       |  ...
         | -   | 9 |   I   |       |       |       |  ...
         -     | . |  ...  |  ...  |  ...  |  ...  |  ...

    Clicking the minus sign on each of the level 2 outlines will collapse
    and hide the data as shown in the next figure. The minus sign changes to
    a plus sign to indicate that the data in the outline is hidden.
	在每个级别2点击减号(-),分级会折叠并且隐藏下一图形中的数据。当减号变为加号时，表明分级显示中的数据被隐藏。

                ------------------------------------------
         1 2 3 |   |   A   |   B   |   C   |   D   |  ...
                ------------------------------------------
          _    | 1 |   A   |       |       |       |  ...
         |     | 2 |   B   |       |       |       |  ...
         | +   | 5 |   E   |       |       |       |  ...
         |     | 6 |   F   |       |       |       |  ...
         | +   | 9 |   I   |       |       |       |  ...
         -     | . |  ...  |  ...  |  ...  |  ...  |  ...

    Clicking on the minus sign on the level 1 outline will collapse the
    remaining rows as follows:
	点击级别1中的减号，分级会按如下方式折叠剩余的行：

                ------------------------------------------
         1 2 3 |   |   A   |   B   |   C   |   D   |  ...
                ------------------------------------------
               | 1 |   A   |       |       |       |  ...
         +     | . |  ...  |  ...  |  ...  |  ...  |  ...

    Grouping in "Excel::Writer::XLSX" is achieved by setting the outline
    level via the "set_row()" and "set_column()" worksheet methods:
	通过"set_row()" 和 "set_column()"方法设置分级显示，"Excel::Writer::XLSX" 能完成数据分组：

        set_row( $row, $height, $format, $hidden, $level, $collapsed )
        set_column( $first_col, $last_col, $width, $format, $hidden, $level, $collapsed )

    The following example sets an outline level of 1 for rows 1 and 2
    (zero-indexed) and columns B to G. The parameters $height and $XF are
    assigned default values since they are undefined:
	下面的例子为1-2行（从0开始索引）和B-G列设置了级别为1的分级显示。参数$height 和 $XF指定了默认值，因为它们是未定义的（undefined）：

        $worksheet->set_row( 1, undef, undef, 0, 1 );
        $worksheet->set_row( 2, undef, undef, 0, 1 );
        $worksheet->set_column( 'B:G', undef, undef, 0, 1 );

    Excel allows up to 7 outline levels. Therefore the $level parameter
    should be in the range "0 <= $level <= 7".
	Excel允许多大7个分级显示。因此，$level参数应该在范围 "0 <= $level <= 7"。

    Rows and columns can be collapsed by setting the $hidden flag for the
    hidden rows/columns and setting the $collapsed flag for the row/column
    that has the collapsed "+" symbol:
	通过为隐藏的行或列设置$hidden标记并且为带有"+"号的行或列设置$collapsed标记来折叠行或列：

        $worksheet->set_row( 1, undef, undef, 1, 1 );
        $worksheet->set_row( 2, undef, undef, 1, 1 );
        $worksheet->set_row( 3, undef, undef, 0, 0, 1 );          # Collapsed flag.

        $worksheet->set_column( 'B:G', undef, undef, 1, 1 );
        $worksheet->set_column( 'H:H', undef, undef, 0, 0, 1 );   # Collapsed flag.

    Note: Setting the $collapsed flag is particularly important for
    compatibility with OpenOffice.org and Gnumeric.
	注意：设置$collapsed标记对于兼容OpenOffice.org 和电子表格特别重要。

    查看"outline.pl"和"outline_collapsed.pl" 程序。

    Some additional outline properties can be set via the
    "outline_settings()" worksheet method, see above.
	一些额外的分级显示属性能通过"outline_settings()"工作表方法设置，查看上面的例子。

DATA VALIDATION IN EXCEL Excel中的数据验证
    Data validation is a feature of Excel which allows you to restrict the
    data that a users enters in a cell and to display help and warning
    messages. It also allows you to restrict input to values in a drop down
    list.
	数据验证是Excel的一种功能，它允许你限制用户在单元格中输入的数据并且显示帮助和警告信息。它也允许你在一个下拉列表中限制输入值。

    A typical use case might be to restrict data in a cell to integer values
    in a certain range, to provide a help message to indicate the required
    value and to issue a warning if the input data doesn't meet the stated
    criteria. In Excel::Writer::XLSX we could do that as follows:
	一个典型的使用实例可能是在一定范围内将单元格中的数据限制为整数。如果输入的数据不符合标准，它会提供帮助信息指定需要的值或发出一个警告。在Excel::Writer::XLSX中我们可以使用如下的方法：

        $worksheet->data_validation('B3',
            {
                validate        => 'integer',
                criteria        => 'between',
                minimum         => 1,
                maximum         => 100,
                input_title     => 'Input an integer:',
                input_message   => 'Between 1 and 100',
                error_message   => 'Sorry, try again.',
            });

   
    The following sections describe how to use the "data_validation()"
    method and its various options.
	下面的章节描述了怎样使用"data_validation()"方法和它的各种选项。

  data_validation( $row, $col, { parameter => 'value', ... } )
    The "data_validation()" method is used to construct an Excel data
    validation.
	"data_validation()"方法用于构建一个Excel数据验证。

    It can be applied to a single cell or a range of cells. You can pass 3
    parameters such as "($row, $col, {...})" or 5 parameters such as
    "($first_row, $first_col, $last_row, $last_col, {...})". You can also
    use "A1" style notation. 例如，:
	它能用于单个单元格或一定范围内的单元格。你可以传递3个参数诸如"($row, $col, {...})"或5个参数诸如 "($first_row, $first_col, $last_row, $last_col, {...})"。你也可以使用A1风格的表示法，例如：

        $worksheet->data_validation( 0, 0,       {...} );
        $worksheet->data_validation( 0, 0, 4, 1, {...} );

        # Which are the same as:

        $worksheet->data_validation( 'A1',       {...} );
        $worksheet->data_validation( 'A1:B5',    {...} );

     
    The last parameter in "data_validation()" must be a hash ref containing
    the parameters that describe the type and style of the data validation.
    The allowable parameters are:
	"data_validation()"中的最后一个参数必须是一个包含描述数据验证的类型和风格参数的散列引用。允许的参数是：
        validate
        criteria
        value | minimum | source
        maximum
        ignore_blank
        dropdown

        input_title
        input_message
        show_input

        error_title
        error_message
        error_type
        show_error

    These parameters are explained in the following sections. Most of the
    parameters are optional, however, you will generally require the three
    main options "validate", "criteria" and "value".
	这些参数在下面的章节中有描述。大多数参数是可选的，然而，你通常需要三个主要选项"validate", "criteria" 和 "value".

        $worksheet->data_validation('B3',
            {
                validate => 'integer',
                criteria => '>',
                value    => 100,
            });

     "data_validation" 方法返回:

         0 成功.
        -1 参数个数不足.
        -2 行或列超界.
        -3 参数或值不正确.

  validate
    This parameter is passed in a hash ref to "data_validation()".
	此参数在散列引用中被传递给"data_validation()"。

	"validate"参数用于设置你想验证的数据类型。该参数总是需要的并且没有默认值。
	允许的值是：

        any
        integer
        decimal
        list
        date
        time
        length
        custom

    *   any is used to specify that the type of data is unrestricted. This
        is the same as not applying a data validation. It is only provided
        for completeness and isn't used very often in the context of
        Excel::Writer::XLSX.
		any用于指定数据类型是无限制的。这与不使用数据验证相同。它只为完整性提供并且不会在Excel::Writer::XLSX内容中经常使用。

    *   integer restricts the cell to integer values. Excel refers to this
        as 'whole number'.
		integer限制单元格的值为整数。Excel将此引用为整数。

            validate => 'integer',
            criteria => '>',
            value    => 100,

    *   decimal限制单元格的值为十进制值。

            validate => 'decimal',
            criteria => '>',
            value    => 38.6,

    *   list restricts the cell to a set of user specified values. These can
        be passed in an array ref or as a cell range (named ranges aren't
        currently supported):
		list限制单元格的值为一列用户指定的值。这些值能在数组引用或单元格范围（目前不支持命名范围）中传递：

            validate => 'list',
            value    => ['open', 'high', 'close'],
            # Or like this:
            value    => 'B1:B3',

        Excel requires that range references are only to cells on the same
        worksheet.
		Excel要求值域引用只是针对同一工作表的单元格的。

    *   date restricts the cell to date values. Dates in Excel are expressed
        as integer values but you can also pass an ISO860 style string as
        used in "write_date_time()". See also "DATES AND TIME IN EXCEL" for
        more information about working with Excel's dates.
		date限制单元格的值为日期。Excel中的日期被计算为整数，但你也可以像"write_date_time()"使用的那样，传递一个ISO860风格的字符串。
            validate => 'date',
            criteria => '>',
            value    => 39653, # 24 July 2008
            # Or like this:
            value    => '2008-07-24T',

    *   time restricts the cell to time values. Times in Excel are expressed
        as decimal values but you can also pass an ISO860 style string as
        used in "write_date_time()". See also "DATES AND TIME IN EXCEL" for
        more information about working with Excel's times.
		time限制单元格的值为时间。Excel中的时间被解释为十进制值，但你也可以像"write_date_time()"使用的那样，传递一个ISO860风格的字符串。

            validate => 'time',
            criteria => '>',
            value    => 0.5, # Noon
            # Or like this:
            value    => 'T12:00:00',

    *   length restricts the cell data based on an integer string length.
        Excel refers to this as 'Text length'.
		length根据一个整数字符串长度限制单元格数据。Excel将该值引用为文本长度。

            validate => 'length',
            criteria => '>',
            value    => 10,

    *   custom restricts the cell based on an external Excel formula that
        returns a "TRUE/FALSE" value.
		custom根据返回“TRUE/FALSE”值的外部Excel公式限制单元格。
            validate => 'custom',
            value    => '=IF(A10>B10,TRUE,FALSE)',

  criteria
    This parameter is passed in a hash ref to "data_validation()".
    该参数在一个散列引用中传递到"data_validation()"。
    The "criteria" parameter is used to set the criteria by which the data
    in the cell is validated. It is almost always required except for the
    "list" and "custom" validate options. It has no default value. Allowable
    values are:
	"criteria"参数用于设置单元格中验证后的数据设置的标准。他几乎总是需要，除了 "list" 和 "custom"验证选项。它没有默认值。允许的值为：

        'between'
        'not between'
        'equal to'                  |  '=='  |  '='
        'not equal to'              |  '!='  |  '<>'
        'greater than'              |  '>'
        'less than'                 |  '<'
        'greater than or equal to'  |  '>='
        'less than or equal to'     |  '<='

    You can either use Excel's textual description strings, in the first
    column above, or the more common symbolic alternatives. The following
    are equivalent:
	你也可以使用Excel的文本描述字符串，在上面的第一列中，或更普通的备选符号。下面的是等价的：

        validate => 'integer',
        criteria => 'greater than',
        value    => 100,

        validate => 'integer',
        criteria => '>',
        value    => 100,

    The "list" and "custom" validate options don't require a "criteria". If
    you specify one it will be ignored.
	"list" 和 "custom"有效性选项不需要以个标准。如果你指定一个它会被忽略。

        validate => 'list',
        value    => ['open', 'high', 'close'],

        validate => 'custom',
        value    => '=IF(A10>B10,TRUE,FALSE)',

  value | minimum | source
    This parameter is passed in a hash ref to "data_validation()".
    该参数在散列引用中被传递给 "data_validation()"。
    The "value" parameter is used to set the limiting value to which the
    "criteria" is applied. It is always required and it has no default
    value. You can also use the synonyms "minimum" or "source" to make the
    validation a little clearer and closer to Excel's description of the
    parameter:
	"value"参数用于对应用了"criteria"的值设置极限值。它总是被需要，并且它没有默认值。你也可以使用同义词 "minimum"或"source"让有效性检验更清晰并且与Excel的参数描述更相近：

        # Use 'value'
        validate => 'integer',
        criteria => '>',
        value    => 100,

        # Use 'minimum'
        validate => 'integer',
        criteria => 'between',
        minimum  => 1,
        maximum  => 100,

        # Use 'source'
        validate => 'list',
        source   => '$B$1:$B$3',

  maximum
    This parameter is passed in a hash ref to "data_validation()".
    该参数在散列引用中被传递给"data_validation()"。
    The "maximum" parameter is used to set the upper limiting value when the
    "criteria" is either 'between' or 'not between':
	
	当"criteria"的值是 'between' 或 'not between'时，"maximum"参数用于设置值的上限。

        validate => 'integer',
        criteria => 'between',
        minimum  => 1,
        maximum  => 100,

  ignore_blank
    This parameter is passed in a hash ref to "data_validation()".
    该参数在散列引用中被传递给"data_validation()"。
    The "ignore_blank" parameter is used to toggle on and off the 'Ignore
    blank' option in the Excel data validation dialog. When the option is on
    the data validation is not applied to blank data in the cell. It is on
    by default.
	
	"ignore_blank"参数用于在Excel的数据有效性检查对话框中开启或关闭'Ignore blank'选项。当该选项开启时，数据有效性检验不会应用到单元格中的空白数据上。默认它是开启的。

        ignore_blank => 0,  # Turn the option off

  dropdown
    This parameter is passed in a hash ref to "data_validation()".
    该参数在散列引用中被传递给"data_validation()"。
    The "dropdown" parameter is used to toggle on and off the 'In-cell
    dropdown' option in the Excel data validation dialog. When the option is
    on a dropdown list will be shown for "list" validations. It is on by
    default.
	
	"dropdown"参数用于在Excel的数据有效性对话框中开启或关闭'In-cell dropdown'选项。当开启该选项时，会因为列表验证而出现下拉列表。默认它是开启的。？

        dropdown => 0,      # Turn the option off

  input_title
    This parameter is passed in a hash ref to "data_validation()".
    该参数在散列引用中被传递给"data_validation()"。
    The "input_title" parameter is used to set the title of the input
    message that is displayed when a cell is entered. It has no default
    value and is only displayed if the input message is displayed. See the
    "input_message" parameter below.
	"input_title"参数用于设置输入信息的标题，它没有默认值，并且只有当输入消息显示时才出现。查看下面的 "input_message" 参数。

        input_title   => 'This is the input title',

    The maximum title length is 32 characters.
	最大的标题长度是32个字符。

  input_message
        该参数在散列引用中被传递给"data_validation()"。


    The "input_message" parameter is used to set the input message that is
    displayed when a cell is entered. It has no default value.
	"input_message"参数用于设置键入单元格时显示的输入消息。它没有默认值。

        validate      => 'integer',
        criteria      => 'between',
        minimum       => 1,
        maximum       => 100,
        input_title   => 'Enter the applied discount:',
        input_message => 'between 1 and 100',

    The message can be split over several lines using newlines, "\n" in
    double quoted strings.
	消息可以使用换行分隔为几行。"\n"在双引号字符串中。

        input_message => "This is\na test.

    The maximum message length is 255 characters.
	消息的最大长度是255个字符。

  show_input
    This parameter is passed in a hash ref to "data_validation()".
    该参数在散列引用中被传递给"data_validation()"。
    The "show_input" parameter is used to toggle on and off the 'Show input
    message when cell is selected' option in the Excel data validation
    dialog. When the option is off an input message is not displayed even if
    it has been set using "input_message". It is on by default.
	
	"show_input"参数用于在Excel的数据有效性检查对话框中开启或关闭'Show input message when cell is selected'选项。当该选项关闭时，输入信息不会显示，即使它设置了"input_message"。默认它是开启的。
        show_input => 0,      # Turn the option off

  error_title
    This parameter is passed in a hash ref to "data_validation()".
	该参数在散列引用中被传递给"data_validation()"。

    The "error_title" parameter is used to set the title of the error
    message that is displayed when the data validation criteria is not met.
    The default error title is 'Microsoft Excel'.

        error_title   => 'Input value is not valid',

    The maximum title length is 32 characters.
	标题的最大长度是32个字符。

  error_message
    This parameter is passed in a hash ref to "data_validation()".
     该参数在散列引用中被传递给"data_validation()"。
    The "error_message" parameter is used to set the error message that is
    displayed when a cell is entered. The default error message is "The
    value you entered is not valid.A user has restricted values that can
    be entered into the cell.".
	
	"error_message" 参数用于设置键入单元格时显示的输入消息。默认的错误消息是"The
    value you entered is not valid."。用户限制了能输入到单元格中的值。


        validate      => 'integer',
        criteria      => 'between',
        minimum       => 1,
        maximum       => 100,
        error_title   => 'Input value is not valid',
        error_message => 'It should be an integer between 1 and 100',

    The message can be split over several lines using newlines, "\n" in
    double quoted strings.
	消息可以使用换行分隔为几行。"\n"在双引号字符串中。
	
	

        input_message => "This is\na test.",

    The maximum message length is 255 characters.
	最大的消息长度值是255个字符。

  error_type
    This parameter is passed in a hash ref to "data_validation()".
	该参数在散列引用中被传递给"data_validation()"。

    The "error_type" parameter is used to specify the type of error dialog
    that is displayed. There are 3 options:
	"error_type"参数用于指定出现的错误对话框的类型。有3个选项：

        'stop'
        'warning'
        'information'

    默认是'stop'.

  show_error
    该参数在散列引用中被传递给"data_validation()".

    The "show_error" parameter is used to toggle on and off the 'Show error
    alert after invalid data is entered' option in the Excel data validation
    dialog. When the option is off an error message is not displayed even if
    it has been set using "error_message". It is on by default.
	
	"show_error"参数用于在Excel的数据有效性检验对话框中开启或关闭'Show error
    alert after invalid data is entered'选项。当该选项关闭时，错误信息不会显示，即使它设置了"error_message"。默认它是开启的。

        show_error => 0,      # Turn the option off

  Data Validation Examples
    Example 1. Limiting input to an integer greater than a fixed value.
	例1.将输入限定为比某一固定值大的整数。

        $worksheet->data_validation('A1',
            {
                validate        => 'integer',
                criteria        => '>',
                value           => 0,
            });

    Example 2. Limiting input to an integer greater than a fixed value where
    the value is referenced from a cell.
   例2.将输入限定为比某一固定值大的整数，该固定值来自单元格引用。
        $worksheet->data_validation('A2',
            {
                validate        => 'integer',
                criteria        => '>',
                value           => '=E3',
            });

    Example 3. Limiting input to a decimal in a fixed range.
	例3.将输入限制为某一固定范围内的十进制值。

        $worksheet->data_validation('A3',
            {
                validate        => 'decimal',
                criteria        => 'between',
                minimum         => 0.1,
                maximum         => 0.5,
            });

    Example 4. Limiting input to a value in a dropdown list.
	例4. 将输入限制为下拉列表中的某个值。

        $worksheet->data_validation('A4',
            {
                validate        => 'list',
                source          => ['open', 'high', 'close'],
            });

    Example 5. Limiting input to a value in a dropdown list where the list
    is specified as a cell range.
	例5.将输入限制为下拉列表中的某个值，该下拉列表由单元格范围指定。

        $worksheet->data_validation('A5',
            {
                validate        => 'list',
                source          => '=$E$4:$G$4',
            });

    Example 6. Limiting input to a date in a fixed range.
	例6.将输入限制为某一固定范围内的日期值。

        $worksheet->data_validation('A6',
            {
                validate        => 'date',
                criteria        => 'between',
                minimum         => '2008-01-01T',
                maximum         => '2008-12-12T',
            });

    Example 7. Displaying a message when the cell is selected.
	例7.当选中单元格时，显示提示消息。

        $worksheet->data_validation('A7',
            {
                validate      => 'integer',
                criteria      => 'between',
                minimum       => 1,
                maximum       => 100,
                input_title   => 'Enter an integer:',
                input_message => 'between 1 and 100',
            });

    查看 "data_validate.pl"程序。

 EXCEL 中的条件格式
    条件格式是Excel的一项功能，允许你根据一定的标准将一个格式应用到一个单元格或一定范围内的单元格中。

	例如，下面的标准用于在"conditional_format.pl"例子中使用红色高亮值大于或等于50的单元格：

        # Write a conditional format over a range.
        $worksheet->conditional_formatting( 'B3:K12',
            {
                type     => 'cell',
                criteria => '>=',
                value    => 50,
                format   => $format1,
            }
        );

  conditional_format( $row, $col, { parameter => 'value', ... } )
	"conditional_format()"方法用于根据用户定义的标准将格式应用到Excel::Writer::XLSX文件中。

	它能被应用到单个单元格中或一定范围内的单元格中。你可以传递3个参数，诸如："($row, $col, {...})" 或5个参数，诸如 "($first_row, $first_col, $last_row, $last_col, {...})".你也可以使用A1表示法，例如：

        $worksheet->conditional_format( 0, 0,       {...} );
        $worksheet->conditional_format( 0, 0, 4, 1, {...} );

        # Which are the same as:

        $worksheet->conditional_format( 'A1',       {...} );
        $worksheet->conditional_format( 'A1:B5',    {...} );

     
     "conditional_format()" 里的最后一个参数必须是一个散列引用，它包含描述数据合法性的类型和风格。主要参数有:
	
	"conditional_format()" 方法中的最后一个参数必须是一个散列引用，该引用包含了描述数据有效性的类型和样式的参数。主要的参数有：

        type
        format
        criteria  #标准
        value
        minimum
        maximum

	用于指定条件格式类型的其它参数在下面的相关章节有显示。

  type
	该参数在散列引用中被传递给"conditional_format()"

	"type"参数用于设置你想应用的条件格式。它总是需要的，并且没有默认值。允许的"type"值和它们的有关参数是：

        Type            Parameters
        ====            ==========
        cell            criteria
                        value
                        minimum
                        maximum

        date            criteria
                        value
                        minimum
                        maximum

        time_period     criteria

        text            criteria
                        value

        average         criteria

        duplicate       (none)

        unique          (none)

        top             criteria
                        value

        bottom          criteria
                        value

        blanks          (none)

        no_blanks       (none)

        errors          (none)

        no_errors       (none)

        2_color_scale   (none)

        3_color_scale   (none)

        data_bar        (none)

        formula         criteria

	所有的格式类型都有"format"参数，看下面。其它类型和参数诸如图标设置会在合适的时间添加。

  type => 'cell'
    This is the most common conditional formatting type. It is used when a
    format is applied to a cell based on a simple criteria. 例如，:
	这是最常用的条件格式类型。根据一个简单的标准，该格式类型在将格式应用到单元格中时被使用。例如：

        $worksheet->conditional_formatting( 'A1',
            {
                type     => 'cell',
                criteria => 'greater than',
                value    => 5,
                format   => $red_format,
            }
        );

    或者使用"between"标准:

        $worksheet->conditional_formatting( 'C1:C4',
            {
                type     => 'cell',
                criteria => 'between',
                minimum  => 20,
                maximum  => 30,
                format   => $green_format,
            }
        );

  criteria # 标准
    The "criteria" parameter is used to set the criteria by which the cell
    data will be evaluated. It has no default value. The most common
    criteria as applied to "{ type => 'cell' }" are:
	"criteria"参数用于设置单元格数据将被计算的标准。它没有默认值。最常用类似"{ type => 'cell' }"的标准有：

        'between'
        'not between'
        'equal to'                  |  '=='  |  '='
        'not equal to'              |  '!='  |  '<>'
        'greater than'              |  '>'
        'less than'                 |  '<'
        'greater than or equal to'  |  '>='
        'less than or equal to'     |  '<='

    You can either use Excel's textual description strings, in the first
    column above, or the more common symbolic alternatives.
	你也可以使用Excel的描述字符串，即上面的第一列，或使用更普通的符号。

    Additional criteria which are specific to other conditional format types
    are shown in the relevant sections below.
	用于指定条件格式类型的其它标准在下面的相关章节有显示。

  value
    
	"value"通常与"criteria"参数一起使用，用于设置将被计算的单元格数据的规则。

        type     => 'cell',
        criteria => '>',
        value    => 5
        format   => $format,

	"value"属性也可以是单元格引用。

        type     => 'cell',
        criteria => '>',
        value    => '$C$1',
        format   => $format,

  format
    The "format" parameter is used to specify the format that will be
    applied to the cell when the conditional formatting criteria is set. The
    format is created using the "add_format()" method in the same way as
    cell formats:
	
	当条件格式标准设置后，"format"参数用于指定将被应用到单元格中的格式。该格式使用与单元格格式一样的"add_format()"方法创建：

        $format = $workbook->add_format( bold => 1, italic => 1 );

        $worksheet->conditional_formatting( 'A1',
            {
                type     => 'cell',
                criteria => '>',
                value    => 5
                format   => $format,
            }
        );

    The conditional format follows the same rules as in Excel: it is
    superimposed over the existing cell format and not all font and border
    properties can be modified. Font properties that can't be modified are
    font name, font size, superscript and subscript. The border property
    that cannot be modified is diagonal borders.
	条件格式允许与Excel同样的规则：它与已经存在的单元格格式重叠，并且不是所有的字体和边框属性能被修改。不能修改的属性有字体名、字体大小、上标和下标。不能被修改的边框属性是斜线边框。

    Excel specifies some default formats to be used with conditional
    formatting. You can replicate them using the following
    Excel::Writer::XLSX formats:
	Excel指定了一些与条件格式一起使用的默认格式。你可以使用下面的Excel::Writer::XLSX的格式复写它们：

        # Light red fill with dark red text.

        my $format1 = $workbook->add_format(
            bg_color => '#FFC7CE',
            color    => '#9C0006',
        );

        # Light yellow fill with dark yellow text.

        my $format2 = $workbook->add_format(
            bg_color => '#FFEB9C',
            color    => '#9C6500',
        );

        # Green fill with dark green text.

        my $format3 = $workbook->add_format(
            bg_color => '#C6EFCE',
            color    => '#006100',
        );

  minimum
    The "minimum" parameter is used to set the lower limiting value when the
    "criteria" is either 'between' or 'not between':
	当标准是'between' 或 'not between'时，"minimum"参数用于设置值的下限：

        validate => 'integer',
        criteria => 'between',
        minimum  => 1,
        maximum  => 100,

  maximum
    The "maximum" parameter is used to set the upper limiting value when the
    "criteria" is either 'between' or 'not between'. See the previous
    example.
	当标准是'between' 或 'not between'时，"maximum"参数用于设置值的上限。查看前一个例子：
  type => 'date'
    The "date" type is the same as "cell" type and uses the same criteria
    and values. However it allows the "value", "minimum" and "maximum"
    properties to be specified in the ISO8601 "yyyy-mm-ddThh:mm:ss.sss" date
    format which is detailed in the "write_date_time()" method.
	"date"类型与"cell"类型相同并使用相同的标准和值。然而，它允许 "value", "minimum" 和 "maximum"属性指定为ISO8601 "yyyy-mm-ddThh:mm:ss.sss"日期格式，它在"write_date_time()"方法上更详细。

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'date',
                criteria => 'greater than',
                value    => '2011-01-01T',
                format   => $format,
            }
        );

  type => 'time_period'
    The "time_period" type is used to specify Excel's "Dates Occurring"
    style conditional format.
	"time_period" 类型用于指定Excel的"Dates Occurring"风格的条件格式。

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'time_period',
                criteria => 'yesterday',
                format   => $format,
            }
        );

    The period is set in the "criteria" and can have one of the following
    values:
    周期在"criteria"中设置，并且可以是如下的值之一：
            criteria => 'yesterday',
            criteria => 'today',
            criteria => 'last 7 days',
            criteria => 'last week',
            criteria => 'this week',
            criteria => 'next week',
            criteria => 'last month',
            criteria => 'this month',
            criteria => 'next month'

  type => 'text'
    The "text" type is used to specify Excel's "Specific Text" style
    conditional format. It is used to do simple string matching using the
    "criteria" and "value" parameters:
	
	"text"类型用于指定Excel的"Specific Text"风格的条件格式。它用于使用"criteria" 和 "value"参数做简单的字符串匹配：

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'text',
                criteria => 'containing',
                value    => 'foo',
                format   => $format,
            }
        );
 
    "criteria"能使用如下的值：
        criteria => 'containing',
        criteria => 'not containing',
        criteria => 'begins with',
        criteria => 'ends with',

	"value"参数可以是一个字符串或单个字符。

  type => 'average'
    The "average" type is used to specify Excel's "Average" style
    conditional format.
	"average"类型用于指定Excel的"Average"风格条件格式。

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'average',
                criteria => 'above',
                format   => $format,
            }
        );

    The type of average for the conditional format range is specified by the
    "criteria":
	条件格式范围的average类型由"criteria"指定：？

        criteria => 'above',
        criteria => 'below',
        criteria => 'equal or above',
        criteria => 'equal or below',
        criteria => '1 std dev above',
        criteria => '1 std dev below',
        criteria => '2 std dev above',
        criteria => '2 std dev below',
        criteria => '3 std dev above',
        criteria => '3 std dev below',

  type => 'duplicate'
    The "duplicate" type is used to highlight duplicate cells in a range:
	"duplicate"类型用于高亮一定范围内完全相同的单元格：

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'duplicate',
                format   => $format,
            }
        );

  type => 'unique'
    The "unique" type is used to highlight unique cells in a range:
	"unique"类型用于高亮一定范围内的唯一单元格：

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'unique',
                format   => $format,
            }
        );

  type => 'top'
    The "top" type is used to specify the top "n" values by number or
    percentage in a range:
	"top"类型用于使用数字或百分比指定单元格的前"n"个值：？

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'top',
                value    => 10,
                format   => $format,
            }
        );

    The "criteria" can be used to indicate that a percentage condition is
    required:
	"criteria"能用于表明需要一个百分比条件：

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'top',
                value    => 10,
                criteria => '%',
                format   => $format,
            }
        );

  type => 'bottom'
    The "bottom" type is used to specify the bottom "n" values by number or
    percentage in a range.
	"bottom"类型用于使用数字或百分比指定单元格的后“n”个值。？

    It takes the same parameters as "top", see above.
	它的参数与“top”一样，看上面。

  type => 'blanks'
    The "blanks" type is used to highlight blank cells in a range:
	"blanks"类型用于在一定范围内高亮空白单元格：

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'blanks',
                format   => $format,
            }
        );

  type => 'no_blanks'
    The "no_blanks" type is used to highlight non blank cells in a range:
	"no_blanks"类型用于在一定范围内高亮非空白单元格：

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'no_blanks',
                format   => $format,
            }
        );

  type => 'errors'
    The "errors" type is used to highlight error cells in a range:
	"errors"类型用于在一定范围内高亮有错误的单元格：

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'errors',
                format   => $format,
            }
        );

  type => 'no_errors'
    The "no_errors" type is used to highlight non error cells in a range:
	"no_errors"类型用于在一定范围内高亮没有错误的单元格：

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'no_errors',
                format   => $format,
            }
        );

  type => '2_color_scale'
    The "2_color_scale" type is used to specify Excel's "2 Color Scale"
    style conditional format.
	"2_color_scale"类型用于指定Excel的"2 Color Scale"风格的条件格式。

        $worksheet->conditional_formatting( 'A1:A12',
            {
                type  => '2_color_scale',
            }
        );

    At the moment only the default colors and properties can be used. These
    will be extended in time.
	现在只能使用默认的颜色和属性。将来会及时进行扩展。

  type => '3_color_scale'
    The "3_color_scale" type is used to specify Excel's "3 Color Scale"
    style conditional format.
	
    "3_color_scale"类型用于指定Excel的"3 Color Scale"风格的条件格式。
        $worksheet->conditional_formatting( 'A1:A12',
            {
                type  => '3_color_scale',
            }
        );

    At the moment only the default colors and properties can be used. These
    will be extended in time.
	
	现在只能使用默认的颜色和属性。将来会及时进行扩展。

  type => 'data_bar'
    The "data_bar" type is used to specify Excel's "Data Bar" style
    conditional format.
    "data_bar"类型用于指定Excel的"data_bar"风格的条件格式。
        $worksheet->conditional_formatting( 'A1:A12',
            {
                type  => 'data_bar',
            }
        );

    At the moment only the default colors and properties can be used. These
    will be extended in time.
    现在只能使用默认的颜色和属性。将来会及时进行扩展。
  type => 'formula'
    The "formula" type is used to specify a conditional format based on a
    user defined formula:
	"formula"类型用于根据用户定义的公式指定一个条件格式：

    $worksheet->conditional_formatting( 'A1:A4', { type => 'formula',
    criteria => '=$A$1 > 5', format => $format, } );

    The formula is specified in the "criteria".
	公式在"criteria"中指定。

  Conditional Formatting Examples
  条件格式例子
    Example 1. Highlight cells greater than or equal to an integer value.
	例1.高亮值大于或等于某个整数值的单元格。

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'cell',
                criteria => 'greater than',
                value    => 5,
                format   => $format,
            }
        );

    Example 2. Highlight cells greater than or equal to a value in a
    reference cell.
	例2.高亮值大于或等于某个值的引用单元格。


        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'cell',
                criteria => 'greater than',
                value    => '$H$1',
                format   => $format,
            }
        );

    Example 3. Highlight cells greater than a certain date:
	例3.高亮其值比某一确定日期大的单元格：

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'date',
                criteria => 'greater than',
                value    => '2011-01-01T',
                format   => $format,
            }
        );

    Example 4. Highlight cells with a date in the last seven days:
	例4.高亮含有最后7天日期的单元格：

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'time_period',
                criteria => 'last 7 days',
                format   => $format,
            }
        );

    Example 5. Highlight cells with strings starting with the letter "b":
	例5.高亮字符串中以字符"b"开头的单元格：

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'text',
                criteria => 'begins with',
                value    => 'b',
                format   => $format,
            }
        );

    Example 6. Highlight cells that are 1 std deviation above the average
    for the range:
	例6.高亮一定范围内，高亮标准差比平均数大1的单元格：

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'average',
                format   => $format,
            }
        );

    Example 7. Highlight duplicate cells in a range:
	例7.高亮一定范围内完全相同的单元格：

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'duplicate',
                format   => $format,
            }
        );

    Example 8. Highlight unique cells in a range.
	例8.高亮在一定范围内唯一的单元格。

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'unique',
                format   => $format,
            }
        );

    Example 9. Highlight the top 10 cells.
	例9.高亮头10个单元格。

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'top',
                value    => 10,
                format   => $format,
            }
        );

    Example 10. Highlight blank cells.
	例10.高亮空白单元格。

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'blanks',
                format   => $format,
            }
        );

    也请查看 "conditional_format.pl"程序。

   Excel中的公式和函数
  Introduction 介绍
  
	下面是Excel和Excel::Writer::XLSX中公式和函数的简明介绍。

	公式是以一个等号开始的字符串：

        '=A1+B1'
        '=AVERAGE(1, 2, 3)'

    The formula can contain numbers, strings, boolean values, cell
    references, cell ranges and functions. Named ranges are not supported.
    Formulas should be written as they appear in Excel, that is cells and
    functions must be in uppercase.
	公式可以包含数字、字符串、布尔值、单元格引用、单元格值域和函数。不支持命名值域。公式应该像在Excel中那样写入，即单元格和函数必须是大写的。

    Cells in Excel are referenced using the A1 notation system where the
    column is designated by a letter and the row by a number. Columns range
    from A to XFD i.e. 0 to 16384, rows range from 1 to 1048576. 
	单元格使用A1表示法系统引用，即用字母标示列，数字标示行。列的范围是从A到XFD或0..16384；行的范围是从1到1048576.例如：

        use Excel::Writer::XLSX::Utility;

        ( $row, $col ) = xl_cell_to_rowcol( 'C2' );    # (1, 2)
        $str = xl_rowcol_to_cell( 1, 2 );              # C2

    The Excel "$" notation in cell references is also supported. This allows
    you to specify whether a row or column is relative or absolute. This
    only has an effect if the cell is copied. The following examples show
    relative and absolute values.
	在Excel里，也支持单元格引用的"$"表示法。这允许你指定行或列是否是相对引用或绝对应用。这只有在复制单元格时有作用。下面的例子显示了相对和绝对值：

        '=A1'   # 列和行是相对的
        '=$A1'  # 列式绝对的，行是相对的
        '=A$1'  # 列是相对的，行是绝对的
        '=$A$1' # 列和行是绝对的

	公式也能引用当前工作簿中其它工作表中的单元格，例如：

        '=Sheet2!A1'
        '=Sheet2!A1:A5'
        '=Sheet2:Sheet3!A1'
        '=Sheet2:Sheet3!A1:A5'
        q{='Test Data'!A1}
        q{='Test Data1:Test Data2'!A1}

    The sheet reference and the cell reference are separated by "!" the
    exclamation mark symbol. If worksheet names contain spaces, commas or
    parentheses then Excel requires that the name is enclosed in single
    quotes as shown in the last two examples above. In order to avoid using
    a lot of escape characters you can use the quote operator "q{}" to
    protect the quotes. See "perlop" in the main Perl documentation. Only
    valid sheet names that have been added using the "add_worksheet()"
    method can be used in formulas. You cannot reference external workbooks.
	工作表引用和单元格引用被感叹号分离。如果工作表名含有空格、逗号或括号，则Excel要求使用单引号将名字引起来，就像上面的最后2个例子一样。  
	
    为避免使用太多转义字符，你可以使用引起操作符"q{}"保护引号。只有使用"add_worksheet()"方法添加的合法工作表名才能被用于公式。你不能引用外部工作簿。
	
    The following table lists the operators that are available in Excel's
    formulas. The majority of the operators are the same as Perl's,
    differences are indicated:
	
	下面的表列出了在Excel公式中可用的操作符。其中大多数操作符与Perl中的操作符相同，不同之处也被指出来了：

        算数操作符:
        =====================
        操作符    含义                      例子
           +      Addition                  1+2
           -      Subtraction               2-1
           *      Multiplication            2*3
           /      Division                  1/4
           ^      Exponentiation            2^3      # 等价于 **
           -      Unary minus               -(1+2)   # 还不支持
           %      Percent (Not modulus)     13%      # 不支持, [1]


        比较操作符:
        =====================
        Operator  Meaning                   Example
            =     Equal to                  A1 =  B1 #等价于==
            <>    Not equal to              A1 <> B1 #等价于!=
            >     Greater than              A1 >  B1
            <     Less than                 A1 <  B1
            >=    Greater than or equal to  A1 >= B1
            <=    Less than or equal to     A1 <= B1


        字符串操作符:
        ================
        Operator  Meaning                   Example
            &     Concatenation             "Hello " & "World!" # [2]


        Reference operators:
        ====================
        Operator  Meaning                   Example
            :     Range operator            A1:A4               # [3]
            ,     Union operator            SUM(1, 2+2, B3)     # [4]


        注意:
		[1]:You can get a percentage with formatting and modulus with MOD().
        [1]: 你可以使用格式化得到一个百分数，使用MOD()得到一个模。
		[2]: 在Perl中与("Hello " . "World!")等价。
		[3]: 该范围等价于单元格 A1, A2, A3和 A4.
        [4]: 逗号与Perl中的列表操作符行为相似。

    The range and comma operators can have different symbols in non-English
    versions of Excel. These will be supported in a later version of
    Excel::Writer::XLSX. European users of Excel take note:
	范围和逗号操作符在非英语版本的Excel中有不同的符号。这些会在以后版本的Excel::Writer::XLSX中支持。欧洲的Excel用户注意：

        $worksheet->write('A1', '=SUM(1; 2; 3)'); # Wrong!!
        $worksheet->write('A1', '=SUM(1, 2, 3)'); # Okay

   
    If your formula doesn't work in Excel::Writer::XLSX try the following:
	如果你的公式在Excel::Writer::XLSX中不起作用，尝试下面的：

        1. Verify that the formula works in Excel (or Gnumeric or OpenOffice.org).
		1.检查公式在Excel中有效
        2. Ensure that cell references and formula names are in uppercase.
		2.确保单元格引用和公式名字是大写的。
        3. Ensure that you are using ':' as the range operator, A1:A4.
		3.确保你使用':'作为范围操作符，A1:A4.
        4. Ensure that you are using ',' as the union operator, SUM(1,2,3).
		4.确保你使用','作为并集操作符，SUM(1,2,3).
        5. Ensure that the function is in the above table.
		5.确保函数是上表中出现的函数。

   

EXAMPLES 例子
    查看 Excel::Writer::XLSX::Examples 获取完整的示例清单.

  例1
    The following example shows some of the basic features of
    Excel::Writer::XLSX.
	下面的例子显示了 Excel::Writer::XLSX的一些基本特征。

        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        # 创建一个叫simple.xlsx的新工作簿并添加一张工作表
        my $workbook  = Excel::Writer::XLSX->new( 'simple.xlsx' );
        my $worksheet = $workbook->add_worksheet();

        # The general syntax is write($row, $column, $token). Note that row and
        # column are zero indexed

        # 写入一些文本
        $worksheet->write( 0, 0, 'Hi Excel!' );


        # Write some numbers
        $worksheet->write( 2, 0, 1 );
        $worksheet->write( 3, 0, 1.00000 );
        $worksheet->write( 4, 0, 2.00001 );
        $worksheet->write( 5, 0, 3.14159 );


        # Write some formulas
        $worksheet->write( 7, 0, '=A3 + A6' );
        $worksheet->write( 8, 0, '=IF(A5>3,"Yes", "No")' );


        # 写入一个超链接
        $worksheet->write( 10, 0, 'http://www.perl.com/' );

  例2
    The following is a general example which demonstrates some features of
    working with multiple worksheets.
	下面是一些普通的例子，它们说明了一些使用多张工作表的特性。

        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        # 创建一个新的Excel工作簿
        my $workbook = Excel::Writer::XLSX->new( 'regions.xlsx' );

        # 添加一些工作表
        my $north = $workbook->add_worksheet( 'North' );
        my $south = $workbook->add_worksheet( 'South' );
        my $east  = $workbook->add_worksheet( 'East' );
        my $west  = $workbook->add_worksheet( 'West' );

        # 添加一个格式
        my $format = $workbook->add_format();
        $format->set_bold();
        $format->set_color( 'blue' );

        #给每张工作表添加一个标题
        for my $worksheet ( $workbook->sheets() ) {
            $worksheet->write( 0, 0, 'Sales', $format );
        }

        # 写入一些数据
        $north->write( 0, 1, 200000 );
        $south->write( 0, 1, 100000 );
        $east->write( 0, 1, 150000 );
        $west->write( 0, 1, 100000 );

        # 设置活动工作表
        $south->activate();

        # 设置第一列的宽度
        $south->set_column( 0, 0, 20 );

        # 设置活动单元格
        $south->set_selection( 0, 1 );

  例3
    Example of how to add conditional formatting to an Excel::Writer::XLSX
    file. The example below highlights cells that have a value greater than
    or equal to 50 in red and cells below that value in green.
	下面是怎样向一个Excel::Writer::XLSX格式的文件添加条件格式的例子。下面的例子使用红色高亮值大于或等于50的单元格，使用绿色高亮值小于50的单元格。

        #!/usr/bin/perl

        use strict;
        use warnings;
        use Excel::Writer::XLSX;

        my $workbook  = Excel::Writer::XLSX->new( 'conditional_format.xlsx' );
        my $worksheet = $workbook->add_worksheet();

		#下面的例子使用红色高亮值大于或等于50的单元格，使用绿色高亮值小于50的单元格
		
        # Light red fill with dark red text.
		#使用暗红色文本进行淡红色填充？
        my $format1 = $workbook->add_format(
            bg_color => '#FFC7CE',
            color    => '#9C0006',

        );

        # Green fill with dark green text.
		#使用暗绿色文本进行绿色填充
        my $format2 = $workbook->add_format(
            bg_color => '#C6EFCE',
            color    => '#006100',

        );

        # Some sample data to run the conditional formatting against.
		#一些简单的用于运行条件格式的数据
        my $data = [
            [ 34, 72,  38, 30, 75, 48, 75, 66, 84, 86 ],
            [ 6,  24,  1,  84, 54, 62, 60, 3,  26, 59 ],
            [ 28, 79,  97, 13, 85, 93, 93, 22, 5,  14 ],
            [ 27, 71,  40, 17, 18, 79, 90, 93, 29, 47 ],
            [ 88, 25,  33, 23, 67, 1,  59, 79, 47, 36 ],
            [ 24, 100, 20, 88, 29, 33, 38, 54, 54, 88 ],
            [ 6,  57,  88, 28, 10, 26, 37, 7,  41, 48 ],
            [ 52, 78,  1,  96, 26, 45, 47, 33, 96, 36 ],
            [ 60, 54,  81, 66, 81, 90, 80, 93, 12, 55 ],
            [ 70, 5,   46, 14, 71, 19, 66, 36, 41, 21 ],
        ];

        my $caption = 'Cells with values >= 50 are in light red. '
          . 'Values < 50 are in light green';

        # 写入数据.
        $worksheet->write( 'A1', $caption );
        $worksheet->write_col( 'B3', $data );

        # 在一定单元格范围内写入条件格式.
        $worksheet->conditional_formatting( 'B3:K12',
            {
                type     => 'cell',
                criteria => '>=',
                value    => 50,
                format   => $format1,
            }
        );

        # 在同一单元格范围内写入另外一个条件格式
        $worksheet->conditional_formatting( 'B3:K12',
            {
                type     => 'cell',
                criteria => '<',
                value    => 50,
                format   => $format2,
            }
        );

  Example 4
	下面程序是使用函数的一个简单的例子：

        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        # 创建一个新的工作簿并添加一张工作表
        my $workbook  = Excel::Writer::XLSX->new( 'stats.xlsx' );
        my $worksheet = $workbook->add_worksheet( 'Test data' );

        # 为第一列设置列宽为
        $worksheet->set_column( 0, 0, 20 );


        # 给标题创建一个格式
        my $format = $workbook->add_format();
        $format->set_bold();


        # 写入抽样数据
        $worksheet->write( 0, 0, 'Sample', $format );
        $worksheet->write( 0, 1, 1 );
        $worksheet->write( 0, 2, 2 );
        $worksheet->write( 0, 3, 3 );
        $worksheet->write( 0, 4, 4 );
        $worksheet->write( 0, 5, 5 );
        $worksheet->write( 0, 6, 6 );
        $worksheet->write( 0, 7, 7 );
        $worksheet->write( 0, 8, 8 );

        $worksheet->write( 1, 0, 'Length', $format );
        $worksheet->write( 1, 1, 25.4 );
        $worksheet->write( 1, 2, 25.4 );
        $worksheet->write( 1, 3, 24.8 );
        $worksheet->write( 1, 4, 25.0 );
        $worksheet->write( 1, 5, 25.3 );
        $worksheet->write( 1, 6, 24.9 );
        $worksheet->write( 1, 7, 25.2 );
        $worksheet->write( 1, 8, 24.8 );

        # 写入一些统计函数
        $worksheet->write( 4, 0, 'Count', $format );
        $worksheet->write( 4, 1, '=COUNT(B1:I1)' );

        $worksheet->write( 5, 0, 'Sum', $format );
        $worksheet->write( 5, 1, '=SUM(B2:I2)' );

        $worksheet->write( 6, 0, 'Average', $format );
        $worksheet->write( 6, 1, '=AVERAGE(B2:I2)' );

        $worksheet->write( 7, 0, 'Min', $format );
        $worksheet->write( 7, 1, '=MIN(B2:I2)' );

        $worksheet->write( 8, 0, 'Max', $format );
        $worksheet->write( 8, 1, '=MAX(B2:I2)' );

        $worksheet->write( 9, 0, 'Standard Deviation', $format );
        $worksheet->write( 9, 1, '=STDEV(B2:I2)' );

        $worksheet->write( 10, 0, 'Kurtosis', $format );
        $worksheet->write( 10, 1, '=KURT(B2:I2)' );

  Example 5
    下面这个例子将一个使用tab字符分隔的叫做“tab.txt”的文件转换为一个叫做"tab.xlsx"的Excel文件。

        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        open( TABFILE, 'tab.txt' ) or die "tab.txt: $!";

        my $workbook  = Excel::Writer::XLSX->new( 'tab.xlsx' );
        my $worksheet = $workbook->add_worksheet();

        # 行和列从0开始索引
        my $row = 0;

        while ( <TABFILE> ) {
            chomp;

            # Split on single tab
            my @fields = split( '\t', $_ );

            my $col = 0;
            for my $token ( @fields ) {
                $worksheet->write( $row, $col, $token );
                $col++;
            }
            $row++;
        }

 	注意：这只是一个说明性的简单的转换程序。转换CSV文件或使用Tab符号分隔或其它任何形式的分隔符分隔的文本文件，我推荐更严密的csv2xls程序，它是Text::CSV_XS模块的一部分。

   在此处查看 examples/csv2xls 链接:
    <http://search.cpan.org/~hmbrand/Text-CSV_XS/MANIFEST>.

  附加的例子
    下面是Excel::Writer::XLSX标准发行版中提供的实例描述文件。它们说明了该模块的不同特征和选项。查看Excel::Writer::XLSX::Examples获取更详细的信息.
        Getting started
        ===============
        a_simple.pl             A simple demo of some of the features.
        bug_report.pl           A template for submitting bug reports.
        demo.pl                 A demo of some of the available features.
        formats.pl              All the available formatting on several worksheets.
        regions.pl              A simple example of multiple worksheets.
        stats.pl                Basic formulas and functions.


        Intermediate
        ============
        autofilter.pl           Examples of worksheet autofilters.
        array_formula.pl        Examples of how to write array formulas.
        cgi.pl                  A simple CGI program.
        chart_area.pl           A demo of area style charts.
        chart_bar.pl            A demo of bar (vertical histogram) style charts.
        chart_column.pl         A demo of column (histogram) style charts.
        chart_line.pl           A demo of line style charts.
        chart_pie.pl            A demo of pie style charts.
        chart_scatter.pl        A demo of scatter style charts.
        chart_stock.pl          A demo of stock style charts.
        colors.pl               A demo of the colour palette and named colours.
        comments1.pl            Add comments to worksheet cells.
        comments2.pl            Add comments with advanced options.
        conditional_format.pl   Add conditional formats to a range of cells.
        data_validate.pl        An example of data validation and dropdown lists.
        date_time.pl            Write dates and times with write_date_time().
        defined_name.pl         Example of how to create defined names.
        diag_border.pl          A simple example of diagonal cell borders.
        filehandle.pl           Examples of working with filehandles.
        headers.pl              Examples of worksheet headers and footers.
        hide_sheet.pl           Simple example of hiding a worksheet.
        hyperlink1.pl           Shows how to create web hyperlinks.
        hyperlink2.pl           Examples of internal and external hyperlinks.
        indent.pl               An example of cell indentation.
        merge1.pl               A simple example of cell merging.
        merge2.pl               A simple example of cell merging with formatting.
        merge3.pl               Add hyperlinks to merged cells.
        merge4.pl               An advanced example of merging with formatting.
        merge5.pl               An advanced example of merging with formatting.
        merge6.pl               An example of merging with Unicode strings.
        mod_perl1.pl            A simple mod_perl 1 program.
        mod_perl2.pl            A simple mod_perl 2 program.
        panes.pl                An examples of how to create panes.
        outline.pl              An example of outlines and grouping.
        outline_collapsed.pl    An example of collapsed outlines.
        protection.pl           Example of cell locking and formula hiding.
        protection.pl           Example of cell locking and formula hiding.
        rich_strings.pl         Example of strings with multiple formats.
        right_to_left.pl        Change default sheet direction to right to left.
        sales.pl                An example of a simple sales spreadsheet.
        stats_ext.pl            Same as stats.pl with external references.
        stocks.pl               Demonstrates conditional formatting.
        tab_colors.pl           Example of how to set worksheet tab colours.
        write_handler1.pl       Example of extending the write() method. Step 1.
        write_handler2.pl       Example of extending the write() method. Step 2.
        write_handler3.pl       Example of extending the write() method. Step 3.
        write_handler4.pl       Example of extending the write() method. Step 4.
        write_to_scalar.pl      Example of writing an Excel file to a Perl scalar.

        Unicode
        =======
        unicode_2022_jp.pl      Japanese: ISO-2022-JP.
        unicode_8859_11.pl      Thai:     ISO-8859_11.
        unicode_8859_7.pl       Greek:    ISO-8859_7.
        unicode_big5.pl         Chinese:  BIG5.
        unicode_cp1251.pl       Russian:  CP1251.
        unicode_cp1256.pl       Arabic:   CP1256.
        unicode_cyrillic.pl     Russian:  Cyrillic.
        unicode_koi8r.pl        Russian:  KOI8-R.
        unicode_polish_utf8.pl  Polish :  UTF8.
        unicode_shift_jis.pl    Japanese: Shift JIS.

LIMITATIONS 限制
    The following limits are imposed by Excel 2007+:

        描述                                 限制        -----------------------------------  ------
        一个字符串中的最大字符数             32,767
        最大的列数                           16,384
        最大的行数                           1,048,576
        工作表名中的最大字符数               31
        页眉/页脚中的最大字符数              254

与Spreadsheet::WriteExcel模块的兼容性
	"Excel::Writer::XLSX"模块是 "Spreadsheet::WriteExcel"模块的替代者

	它支持 Spreadsheet::WriteExcel所有的特性，注意下面微小的不同：

        工作簿方法                  支持
        ================            ======
        new()                       Yes
        add_worksheet()             Yes
        add_format()                Yes
        add_chart()                 Yes
        close()                     Yes
        set_properties()            Yes
        define_name()               Yes
        set_tempdir()               Yes
        set_custom_color()          Yes
        sheets()                    Yes
        set_1904()                  Yes
        set_optimization()          Yes. Spreadsheet::WriteExcel中不需要.
        add_chart_ext()             Not supported.Excel::Writer::XLSX中不是必须的
        compatibility_mode()        Deprecated. Excel::Writer::XLSX中不是必须的
        set_codepage()              Deprecated. Excel::Writer::XLSX中不是必须的


        Worksheet Methods           Support
        =================           =======
        write()                     Yes
        write_number()              Yes
        write_string()              Yes
        write_rich_string()         Yes. Spreadsheet::WriteExcel中没有该方法.
        write_blank()               Yes
        write_row()                 Yes
        write_col()                 Yes
        write_date_time()           Yes
        write_url()                 Yes
        write_formula()             Yes
        write_array_formula()       Yes.Spreadsheet::WriteExcel中没有该方法.
        keep_leading_zeros()        Yes
        write_comment()             Yes
        show_comments()             Yes
        set_comments_author()       Yes
        add_write_handler()         Yes
        insert_image()              Yes/Partial, 查看文档.
        insert_chart()              Yes
        data_validation()           Yes
        conditional_format()        Yes. Spreadsheet::WriteExcel中没有该方法.
        get_name()                  Yes
        activate()                  Yes
        select()                    Yes
        hide()                      Yes
        set_first_sheet()           Yes
        protect()                   Yes
        set_selection()             Yes
        set_row()                   Yes.
        set_column()                Yes.
        outline_settings()          Yes
        freeze_panes()              Yes
        split_panes()               Yes
        merge_range()               Yes
        merge_range_type()          Yes. Spreadsheet::WriteExcel中没有该方法。
        set_zoom()                  Yes
        right_to_left()             Yes
        hide_zero()                 Yes
        set_tab_color()             Yes
        autofilter()                Yes
        filter_column()             Yes
        filter_column_list()        Yes. Spreadsheet::WriteExcel中没有该方法.
        write_utf16be_string()      不推荐使用. 使用 Perl utf8字符串代替.
        write_utf16le_string()      不推荐使用. 使用 Perl utf8字符串代替.
        store_formula()             不推荐使用. 查看文档.
        repeat_formula()            不推荐使用. 查看文档.
        write_url_range()           Not supported. Excel::Writer::XLSX中不是必须的

        页面设置方法                支持
        ===================         =======
        set_landscape()             Yes
        set_portrait()              Yes
        set_page_view()             Yes
        set_paper()                 Yes
        center_horizontally()       Yes
        center_vertically()         Yes
        set_margins()               Yes
        set_header()                Yes
        set_footer()                Yes
        repeat_rows()               Yes
        repeat_columns()            Yes
        hide_gridlines()            Yes
        print_row_col_headers()     Yes
        print_area()                Yes
        print_across()              Yes
        fit_to_pages()              Yes
        set_start_page()            Yes
        set_print_scale()           Yes
        set_h_pagebreaks()          Yes
        set_v_pagebreaks()          Yes

        格式方法                    支持
        ==============              =======
        set_font()                  Yes
        set_size()                  Yes
        set_color()                 Yes
        set_bold()                  Yes
        set_italic()                Yes
        set_underline()             Yes
        set_font_strikeout()        Yes
        set_font_script()           Yes
        set_font_outline()          Yes
        set_font_shadow()           Yes
        set_num_format()            Yes
        set_locked()                Yes
        set_hidden()                Yes
        set_align()                 Yes
        set_rotation()              Yes
        set_text_wrap()             Yes
        set_text_justlast()         Yes
        set_center_across()         Yes
        set_indent()                Yes
        set_shrink()                Yes
        set_pattern()               Yes
        set_bg_color()              Yes
        set_fg_color()              Yes
        set_border()                Yes
        set_bottom()                Yes
        set_top()                   Yes
        set_left()                  Yes
        set_right()                 Yes
        set_border_color()          Yes
        set_bottom_color()          Yes
        set_top_color()             Yes
        set_left_color()            Yes
        set_right_color()           Yes

REQUIREMENTS  要求
    <http://search.cpan.org/search?dist=Archive-Zip/>.

    Perl 5.10.0.

SPEED AND MEMORY USAGE  速度和内存使用

	"Spreadsheet::WriteExcel"用于优化速度并减少内存使用。这样的设计目标意味着实现许多用户要求的诸如格式化和单独地写入数据功能并不容易。

     因此，"Excel::Writer::XLSX"采取不同的设计方法并且在内存中存入更多的数据，以至于功能更复杂。这样的结果就是Excel::Writer::XLSX 比Spreadsheet::WriteExcel 慢了50%，并且显著地使用更多内存。当你增加更多的行和列范围时，可能因为创建大的文件而用光内存。对于Spreadsheet::WriteExcel这从来都不是问题。
   
	使用工作簿的"set_optimization()"方法，这种内存使用几乎能完全减小：

        $workbook->set_optimization();

 
	这样做的代价就是你不能使用任何操作单元格的新功能，写入数据后，该优化选项被打开。


DIAGNOSTICS  诊断
    Filename required by Excel::Writer::XLSX->new()
       必须给构造函数一个文件名。

    Can't open filename. It may be in use or protected.
       
		不能打开文件写入。你要写入的那个文件夹可能被写保护或文件正被其它程序使用。

    Can't call method "XXX" on an undefined value at someprogram.pl.
      
    在Windows上这常常是由于你正尝试创建的文件与一个已经被Excel打开并锁住的文件冲突了。
    The file you are trying to open 'file.xls' is in a different format than
    specified by the file extension.

		当你创建一个XLSX文件，但是给它一个xls后缀时，会出现该警告。

WRITING EXCEL FILES
    
    根据你的需求，背景和一般的感觉，你可能更喜欢下面的方法之一将数据写入Excel：
    *   Spreadsheet::WriteExcel

 
		它是Excel::Writer::XLSX 的先驱，并使用同样的接口。它生成xls格式的文件用于Excel 97-2003版本。这些文件仍然能被Excel2007读取但是在支持的行和列数量上有限制。

    *   Win32::OLE module and office automation

         这需要一个Windows平台并安装一份Excel拷贝。这是与Excel交互的最强大最完全的方法。
    *   CSV, comma separated variables or text

        如果文件扩展名是csv，Excel打开后会自动转换该格式。生成一个CSV文件并不像看起来的那么容易。查看DBD::RAM, DBD::CSV, Text::xSV 和 Text::CSV_XS 模块。
    *   DBI with DBD::ADO or DBD::ODBC

        Excel文件含有内部索引表，允许它们表现的像一个数据库。使用标准的Perl数据库模块之一，你可以将一个Excel文件当作数据库连接。

    *   DBD::Excel

         
		你也可以通过 DBD::Excel 模块使用标准的DBI接口访问Spreadsheet::WriteExcel 。
		查看  <http://search.cpan.org/dist/DBD-Excel>
    *   Spreadsheet::WriteExcelXML
       
        该模块允许你使用与 Spreadsheet::WriteExcel同样的接口来创建Excel XML 文件。
		查看 <http://search.cpan.org/dist/Spreadsheet-WriteExcelXML>.
		
    *   Excel::Template

        该模块允许你从在某种意义上与 HTML::Template相似的XML 模版上创建文件。
		查看<http://search.cpan.org/dist/Excel-Template/>.
        

    *   Spreadsheet::WriteExcel::FromXML

       
		该模块允许你使用Spreadsheet::WriteExcel 作为后台将一个简单的XML文件转换为Excel文件。
		XML的格式由支持的DTD来定义。
		查看<http://search.cpan.org/dist/Spreadsheet-WriteExcel-FromXML>.

    *   Spreadsheet::WriteExcel::Simple

       它提供了对Spreadsheet::WriteExcel更加简单的接口。
        <http://search.cpan.org/dist/Spreadsheet-WriteExcel-Simple>.

    *   Spreadsheet::WriteExcel::FromDB

  	 它对于从DB表中创建Excel文件很有用。
	 <http://search.cpan.org/dist/Spreadsheet-WriteExcel-FromDB>.

    *   HTML tables

        This is an easy way of adding formatting via a text based format.
		通过一个基于文本的格式添加格式很容易。

    *   
READING EXCEL FILES  读取Excel文件
	从Excel中读取数据，尝试：

    *   Spreadsheet::ParseExcel

        
        它使用OLE::Storage-Lite模块从Excel中提取数据。
		查看<http://search.cpan.org/dist/Spreadsheet-ParseExcel>.
    *   Spreadsheet::ParseExcel_XLHTML


    *   XML::Excel
         使用Spreadsheet::ParseExcel将Excel文件转换为XML文件
        查看那<http://search.cpan.org/dist/XML-Excel>.

    *   
