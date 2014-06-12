----------------------------------------------------------------------------
 ����: СPerl
Email:sxw2k@sina.com
----------------------------------------------------------------------------



NAME     ����
    Excel::Writer::XLSX - ��Excel2007+XLSX��ʽ����һ�����ļ�.

VERSION  �汾
    ���ĵ�����2013��11�·�����Excel::Writer::XLSX 0.75�汾��
SYNOPSIS ��Ҫ
    ��perl.xlsx�ĵ�һ����������д���ַ�������ʽ�����ַ��������ֺ͹�ʽ��
        use Excel::Writer::XLSX;

        # �½�excel������
        my $workbook = Excel::Writer::XLSX->new( 'perl.xlsx' );

        # ����һ��������
        $worksheet = $workbook->add_worksheet();

        #  ���Ӳ�����һ����ʽ        
        $format = $workbook->add_format();#����һ�ָ�ʽ
        $format->set_bold();              #���ô���
        $format->set_color( 'red' );      #������ɫ
        $format->set_align( 'center' );   #���ö��뷽ʽ���˴�Ϊ���У�

		#д��һ����ʽ���ͷǸ�ʽ�����ַ�����ʹ�����б�ʾ����
        $col = $row = 0;                  #�����к��е�λ��
        $worksheet->write( $row, $col, 'Hi Excel!', $format );
        $worksheet->write( 1, $col, 'Hi Excel!' );

        #ʹ��A1��ʾ��д��һ�����ֺ͹�ʽ
        $worksheet->write( 'A3', 1.2345 );         #�ڵ����е�һ��д��һ������
        $worksheet->write( 'A4', '=SIN(PI()/4)' ); #�ڵ����е�һ��д��һ����ʽ

DESCRIPTION ˵��
    The "Excel::Writer::XLSX" ģ�����Ա���������Excel2007+ XLSX��ʽ���ļ���

    XLSX��ʽ��Excel 2007 ���Ժ��汾ʹ�õĹٷ�����XML��OOXML����ʽ
    ���ڹ����������Ӷ��Ź���������ʽ���Ա�Ӧ�õ���Ԫ���С����԰����֣����֣��͹�ʽд�뵥Ԫ����

   ��ģ��Ŀǰ�����ܱ�������һ���Ѿ����ڵ�EXCEl XLSX�ļ���д�����ݡ�

Excel::Writer::XLSX and Spreadsheet::WriteExcel
   
	Excel::Writer::XLSXʹ�ú�Spreadsheet::WriteExcelģ����ͬ�Ľӿ������ɶ�����XLS��ʽ��Excel�ļ�

    Excel::Writer::XLSX ֧������Spreadsheet::WriteExcel�е����ԣ�����ĳЩ�����¹��ܸ�ǿ�����鿴 "Compatibility with Spreadsheet::WriteExcel".��ȡ����ϸ�ڡ�

    
    XLSL��ʽ����XLS��ʽ��Ҫ����������������һ�������������ɸ����������к��С�
QUICK START ��������
     Excel::Writer::XLSX��ͼ�����ܵ��ṩExcel�Ĺ��ܽӿڡ����ˣ��кܶ����ӿ��йص��ĵ�����һ�ۺ��ѿ�����Щ��Ҫ����Щ����Ҫ�����Զ���������Щ��ϲ������װ�˼��豸���ٶ�˵�������ˣ��˴������ּ򵥵ķ�ʽ��
	0���½�һ��Excel����
    1��ʹ��"new()"��������һ���µ�Excel ������
    2��ʹ��"add_worksheet()"�������¹���������һ��������
	3��ʹ��"write()"��������������д������
    ����������

        use Excel::Writer::XLSX;                                   # Step 0

        my $workbook = Excel::Writer::XLSX->new( 'perl.xlsx' );    # Step 1
        $worksheet = $workbook->add_worksheet();                   # Step 2
        $worksheet->write( 'A1', 'Hi Excel!' );                    # Step 3
    
	���ᴴ��һ������perl.xlsx��Excel�ļ�������ֻ��һ�Ź������������ص�Ԫ��������'Hi Excel'�ı���
    
   
����������
       Excel::Writer::XLSXģ��Ϊ�½���Excel�������ṩ�����������Ľӿڡ������ķ�������ͨ��һ���½��Ĺ�������������.

        new()                       #�½�
        add_worksheet()             #���ӹ�����
        add_format()                #���Ӹ�ʽ
        add_chart()                 #����ͼ��
        close()                     #�رչ�����
        set_properties()			#��������
        define_name()				#��������
        set_tempdir()				#������ʱ�ļ���
        set_custom_color()			#�����Զ�����ɫ
        sheets()					#������
        set_1904()                  #���ü�Ԫ��ʼ��
        set_optimization()          #�����Ż�

  new()   
    ʹ��"new()"���췽������һ���µ�Excel���������÷�������һ���ļ������ļ�������Ϊ���������������Ӹ���һ���ļ���������һ���µ�Excel�ļ���
	
        my $workbook  = Excel::Writer::XLSX->new( 'filename.xlsx' );
        my $worksheet = $workbook->add_worksheet();
        $worksheet->write( 0, 0, 'Hi Excel!' );

    ������ʹ���ļ�����Ϊnew()�����������������ӣ�
	
        my $workbook1 = Excel::Writer::XLSX->new( $filename );
        my $workbook2 = Excel::Writer::XLSX->new( '/tmp/filename.xlsx' );
        my $workbook3 = Excel::Writer::XLSX->new( "c:\\tmp\\filename.xlsx" );#Windows
        my $workbook4 = Excel::Writer::XLSX->new( 'c:\tmp\filename.xlsx' );

    ������������˵��������ͨ��ת��Ŀ¼�ָ���"\"��ʹ�õ����ű�ֵ֤�����ڲ�����DOS�ϻ�Windows�Ͻ���Excel�ļ���
	
    �����Ƽ��ļ���ʹ��".xlsx"������".xls"��׺����Ϊ������ʹ��XLSX��ʽ���ļ�ʱ�ᷢ�����档
  
  
	"new()"���캯����������һ��Excel::Writer::XLSX������������ʹ���������������ӹ��������洢���ݡ�	Ӧ��ע�����ǣ�����û���ر�Ҫ��ʹ��"my"���������������¹����������������򣬲��ң��ڴ����������£�����֤�˹�����������ʽ�ص���"close()����"���ܱ���ȷ�عرա�
	
    �����ļ����ܱ������������ļ�Ȩ�޻�����һЩԭ����"new"�᷵��"undef"�����ˣ��ڼ���֮ǰ����"new"�ķ���ֵ�Ǹ���ϰ�ߡ�ͨ�������������ļ�����������Perl����$!�ͻᱻ���ã�

        my $workbook = Excel::Writer::XLSX->new( 'protected.xlsx' );
        die "Problems creating new Excel file: $!" unless defined $workbook;
		
	��Ҳ���Դ���һ���Ϸ����ļ�������"new()"���캯����������һ��CGI��������������������

        binmode( STDOUT );
        my $workbook = Excel::Writer::XLSX->new( \*STDOUT );

    
    ����CGI��������Ҳ����ʹ���ر���Perl�ļ��� '-'�������������ض��򵽱�׼������
        my $workbook = Excel::Writer::XLSX->new( '-' );

   ���Բ鿴�����е�cgi.pl

    Ȼ��������������������"mod_perl"�����в������ã���������һЩ���������飺
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

    ���鿴mod_perl1.pl" �� "mod_perl2.pl"
  
  
	��������ͨ��socket ȥ streamһ��Excel�ļ�����������һ��Excel�ļ�����һ��������
	��ô�ļ������������á�
	���磬�����ǰ�Excel�ļ�д��������һ�ַ�����
        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        open my $fh, '>', \my $str or die "Failed to open filehandle: $!";

        my $workbook  = Excel::Writer::XLSX->new( $fh );
        my $worksheet = $workbook->add_worksheet();

        $worksheet->write( 0, 0, 'Hi Excel!' );

        $workbook->close();

        #  Excel�ļ�������$str�С��ǵ��ڴ�ӡ$str֮ǰbinmode()�����ļ�����
        binmode STDOUT;
        print $str;

    ���鿴"write_to_scalar.pl" �� "filehandle.pl" 

    ע��binmode()��Ҫ����һ��Excel�ļ��ɶ������������ɡ����ˣ�������ʹ��һ���ļ�������Ӧ�ñ�֤�ڴ��ݸ�"new()"����֮ǰ����binmode().������ʹ��Windows����������ƽ̨���㶼Ӧ����������
    
    ������ʹ���ļ��������ļ��������㲻�õ���"binmode()".��Excel::Writer::XLSX���ļ���ת��Ϊ�ļ�����ʱ���������ڲ�ִ��binmode().
---------------------------------------------------------------------------------------------------------	
  add_worksheet( $sheetname )   #����Ϊ��������
   
	����һ��������Ӧ�ñ����ӵ��������С������������ڽ�����д�뵥Ԫ����
        $worksheet1 = $workbook->add_worksheet();               # Sheet1
        $worksheet2 = $workbook->add_worksheet( 'Foglio2' );    # Foglio2
        $worksheet3 = $workbook->add_worksheet( 'Data' );       # Data
        $worksheet4 = $workbook->add_worksheet();               # Sheet4


    ����û��ָ��$sheetname��Ĭ�ϻ�ʹ��Sheet1,Sheet2....
    
	
    �������������ǺϷ���Excel���������������ܰ��������ַ�,"[ ] : * ? / \" ,���ҳ��ȱ���С��32���ַ������⣬�㲻����һ�����ϵĹ�������ʹ��ͬһ���ļ���������Сд���е��ļ�����
	
----------------------------------------------------------------------------------------------
  add_format( %properties )   #���Ӹ�ʽ
    
	    "add_format()"�������Ա����������µĸ�ʽ�����������Ա����ڽ���ʽӦ�õ���Ԫ���С��������ڴ���ʱͨ����������ֵ�Ĺ�ϣ�������Ի�֮��ͨ���������ö������ԣ�
		
	$format1 = $workbook->add_format( %props );    # �ڴ���ʱ��������
    $format2 = $workbook->add_format();            # �������ٶ�������

    ���鿴����Ԫ����ʽ�����½ڻ�ȡ��ϸ��Ϣ
---------------------------------------------------------------------------------------------------------
  add_chart( %properties )  #����ͼ��
    	�÷��������½�һ��ͼ����Ϊһ�������Ĺ�������Ĭ�ϣ�������Ϊһ����Ƕ���Ķ�����ͨ��"insert_chart()"�������������뵽�������С�
	
    my $chart = $workbook->add_chart( type => 'column' );

    ���Կ�������Ϊ:

        type     (required)#������ѡ��
        subtype  (optional)#ͼ��������(��ѡ)
        name     (optional)
        embedded (optional)

    *   "type" ����

        ���Ǳ����Ĳ������������˽���������ͼ�������͡�
         my $chart = $workbook->add_chart( type => 'line' );

        ���õ�type��������:

            area		#����ͼ
            bar			#����ͼ
            column		#����ͼ
            line		#��ͼ
            pie         #��ͼ
            scatter     #ɢ��ͼ
            stock		#����ͼ

    *   "subtype"       #ͼ�����ͣ������ͣ�

        
        ��������Ҫʱ����һ��ͼ����������
            my $chart = $workbook->add_chart( type => 'bar', subtype => 'stacked' );

        
        Ŀǰֻ������ͼ������ͼ֧��������(stacked and percent_stacked)
		
    *   "name"

        Ϊͼ���������֡����������ǿ�ѡ�ģ�����������֧�֣�����Ĭ��Ϊ"Chart1,Chart2....Chartn".
		ͼ���������ǺϷ��ı�������"add_worksheet()"����һ����"name"���Կ�����Ƕ�׵�ͼ����ʡ�ԡ�
            my $chart = $workbook->add_chart( type => 'line', name => 'Results Chart' );

    *   "embedded"

        ָ��ͼ��������ͨ��"insert_chart()"�������������뵽�������С�����û������������־�ͳ��Բ���ͼ���������ִ�����
            my $chart = $workbook->add_chart( type => 'line', embedded => 1 );

            # Configure the chart.
            ...

            # ��ͼ�����뵽��������
            $worksheet->insert_chart( 'E2', $chart );

	�鿴Excel::Writer::XLSX::Chart��ȡ����ϸ�Ĺ����ڴ�������������ͼ����������Ϣ��Ҳ�ɲ鿴chart_*.pl������


  close()
        һ���أ������ĳ������������������󳬳�������ʱ������Excel�ļ��ᱻ�Զ��رա�Ȼ����������ʹ��close()������ʽ�عر�Excel�ļ���
        $workbook->close();#��ʾ�عر�Excel

       ����Excel�ļ������ڶ���ִ��һЩ�ⲿ�������縴�ơ���ȡ��С���߰�����Ϊ�����ʼ��ĸ���֮ǰ�رգ���Ҫ��ʽ����close()����
        ���⣬"close()"��������ֹPerl�������������Դ�����˳���������������������͸�ʽ�����������������������棺
        ����"my()"û�б���������ʹ��"new()"�����Ĺ���������������
        �������������е���"new()", "add_worksheet()" ���� "add_format()"������
		
  
    ԭ����Excel::Writer::XLSX����Perl��"DESTROY"�������ض�˳�򴥷�destructor���������������������������͸�ʽ�������Ǵʷ�������������ӵ�в�ͬ�Ĵʷ�������ʱ��ǰ�������������ᷢ����
    һ���أ������㴴��һ��0�ֽڵ��ļ������㲻�ܽ���һ���ļ�������Ҫ����"close()"������
    
	  "close()"�ķ���ֵ��perl�ر�ʹ��"new()"�����������ļ��ķ���ֵһ�������������Գ��淽ʽ��������
        $workbook->close() or die "Error closing file: $!";

  set_properties()
  
	"set_properties" �����ɱ���������ͨ��"Excel::Writer::XLSX"ģ�鴴����Excel�ļ����ĵ����ԡ�
    ����ʹ��Excel�е�"�칫��ť" ->"׼��"->"����"ѡ��ʱ�����Կ�����Щ���ԡ�

	����ֵӦ���Թ�ϣ��ʽ���ݣ����£�

        $workbook->set_properties(
            title    => 'This is an example spreadsheet',
            author   => 'John McNamara',
            comments => 'Created with Perl and Excel::Writer::XLSX',
        );

    ���Ա����õ�������:

        title    #����
        subject	 #����	
        author	 #����
        manager  #����
        company  #��˾
        category #����
        keywords #�ؼ���
        comments #ע�� 
        status	 #״̬

    ���鿴"properties.pl" ������

  define_name()
    
    �÷��������ڶ���һ�����֣����ܱ����ڱ�ʾ�������е�һ��ֵ��һ�������ĵ�Ԫ�񣬻�һ����Χ�ڵĵ�Ԫ��
	���磺����һ�� global/workbook ��:

        # Global/workbook names.
        $workbook->define_name( 'Exchange_rate', '=0.96' );
        $workbook->define_name( 'Sales',         '=Sheet1!$G$1:$H$10' );

    Ҳ����ʹ���﷨"sheetname!definedname"������֮ǰ���ϱ���������һ�� local/worksheet:
        # Local/worksheet name.
        $workbook->define_name( 'Sheet2!Sales',  '=Sheet2!$G$1:$G$10' );

    ���������������пո��������ַ�������������Excel��һ�����õ����Ž�������������
        $workbook->define_name( "'New Data'!Sales",  '=Sheet2!$G$1:$G$10' );

    �鿴 defined_name.pl ������

  set_tempdir()
   
     "Excel::Writer::XLSX"����װ�������Ĺ�����֮ǰ�������ݴ洢����ʱ�ļ���
     "File::Temp"ģ�����ڴ�����Щ��ʱ�ļ���File::Tempģ��ʹ��"File::Spec"Ϊ��Щ��ʱ�ļ�ָ��һ�����ʵ�λ�ã�����"/tmp"��"c:\windows\temp".�����԰������ķ����ҳ���ϵͳ���ĸ�Ŀ¼��ʹ���ˣ�
        perl -MFile::Spec -le "print File::Spec->tmpdir()
    ����Ĭ�ϵ���ʱ�ļ�Ŀ¼����ʹ�ã�������ʹ��"set_tempdir()"����ָ��һ���ɹ�ѡ����λ�ã�
        $workbook->set_tempdir( '/tmp/writeexcel' );
        $workbook->set_tempdir( 'c:\windows\temp\writeexcel' );

    ���ڴ�����ʱ�ļ���Ŀ¼�����ȴ��ڣ���set_temp()�����������½�һ��Ŀ¼��
    һ��Ǳ��������һЩWindowsϵͳ��������ʱ�ļ�����������Ϊ��Լ800��������ζ�ţ�һ���ڸ���ϵͳ�����еĵ������򽫻ᱻ���ƴ����ܹ�800���������͹�����������������Ҫ�����������ж����ǲ�����������������������
  set_custom_color( $index, $red, $green, $blue )
   #�����Զ�����ɫֵ
    "set_custom_color()"����������ʹ�ø����ʵ���ɫ��������֮һ���ڽ���ɫֵ��
	$index��ֵӦ����8..63֮�䣬�鿴see "COLOURS IN EXCEL".

	Ĭ�ϵ�������ɫʹ������������

         8   =>   black
         9   =>   white
        10   =>   red
        11   =>   lime       #�̻�ɫ
        12   =>   blue
        13   =>   yellow
        14   =>   magenta    #����ɫ
        15   =>   cyan       #����ɫ
        16   =>   brown
        17   =>   green
        18   =>   navy      #����ɫ
        20   =>   purple    #��ɫ
        22   =>   silver    #��ɫ
        23   =>   gray      #��ɫ
        33   =>   pink      #�ۺ�ɫ
        53   =>   orange

    ʹ������RGB(red green blue)�ɷ���������ɫ�� $red,$green �� $blue��ֵ��Χ������0..255֮�䡣
	��������Excel��ʹ��"����"->ѡ��->��ɫ->�޸�"�Ի���������Ҫ����ɫ��
	
	"set_custom_color()"��������������ʹ��HTML������ʮ������ֵ��

        $workbook->set_custom_color( 40, 255,  102,  0 );       # Orange
        $workbook->set_custom_color( 40, 0xFF, 0x66, 0x00 );    # Same thing
        $workbook->set_custom_color( 40, '#FF6600' );           # Same thing

        my $font = $workbook->add_format( color => 40 );        # Modified colour

   
    "set_custom_color()"�����ķ���ֵ�Ǳ��޸ĵ���ɫ��������
        my $ferrari = $workbook->set_custom_color( 40, 216, 12, 12 );

        my $format = $workbook->add_format(
            bg_color => $ferrari,
            pattern  => 1,
            border   => 1
        );

    ע�⣬��XLSX��ʽ�У���ɫ��ɫ�岻ȷ�о���Ϊ53�ִ�ɫ��Excel::Writer::XLSXģ�������Ժ��Ľ׶���չ��֧���µģ������޵ĵ�ɫ�塣
  sheets( 0, 1, ... )
   
	"sheets()"��������һ���������й��������б������б���Ƭ

	����û�д��ݲ�����sheet()�������򷵻ع������е����й�����������������һ�������������ظ��������⽫�����á�

        for $worksheet ( $workbook->sheets() ) {
            print $worksheet->get_name();
        }

  
	������ָ��һ���б���Ƭ����һ��������������������
	
        $worksheet = $workbook->sheets( 0 );
        $worksheet->write( 'A1', 'Hello' );

    ������Ϊ"sheets()"�ķ���ֵ��һ���Թ��������������ã������Խ�����������дΪ��
        $workbook->sheets( 0 )->write( 'A1', 'Hello' );

	
	���������ӷ���һ���������еĵ�һ��������һ����������

        for $worksheet ( $workbook->sheets( 0, -1 ) ) {
            # Do something
        }


  set_1904()
   Excel�����ݴ洢Ϊʵ�������������ִ洢���¼�Ԫ��������������С�����ִ洢һ���İٷֱȡ��¼�Ԫ������1900��1904��Windows�ϵ�Excelʹ��1900��Mac�ϵ�Excelʹ��1904.Ȼ�����κ�ƽ̨�ϵ�Excel������ϵͳ֮���Զ�ת����
   Excel::Writer::XLSXĬ��ʹ��1900��ʽ�洢���ݡ����������ı����������Ե���"set_1904()"����������������1900������0������1904������1.


  set_optimization()
  
	"set_optimization()" �������ڴ���Excel::Writer::XLSXģ���е��Ż�������Ŀǰֻ��һ�������ڴ�ʹ�õ��Ż�������

        $workbook->set_optimization();


    ע�⣬�򿪴��Ż������󣬵�ͨ��"write_*()"�����е�����֮һ������������һ����Ԫ������һ�����ݱ�д��Ȼ����ɾ������Ϊһ���Ż�������������������Ӧ������������˳��д�롣��
    
	�÷����������κε���"add_worksheet()"����֮ǰ�����á�

WORKSHEET METHODS ����������
	
	ͨ�����ù����������е�"add_worksheet()"��������һ���µĹ�����:

        $worksheet1 = $workbook->add_worksheet();
        $worksheet2 = $workbook->add_worksheet();

	�����ķ�������һ���µ�worksheet�ǿ��õģ�

        write()
        write_number()
        write_string()
        write_rich_string()
        keep_leading_zeros()    #����ǰ��0
        write_blank()
        write_row()
        write_col()
        write_date_time()
        write_url()              #д��url
        write_url_range()
        write_formula()#д�빫ʽ
        write_comment()#д��ע��
        show_comments()#��ʽע��
        set_comments_author()
        add_write_handler()
        insert_image()#����ͼ��
        insert_chart()#����ͼ��
        data_validation()#���ݼ���
        conditional_format()
        get_name()
        activate()#����
        select()
        hide()
        set_first_sheet()
        protect()
        set_selection()
        set_row()
        set_column()
        outline_settings()
        freeze_panes()          #���ᴰ��
        split_panes()		    #�ָ��
        merge_range()			#�ϲ�ֵ��
        merge_range_type()
        set_zoom()
        right_to_left()
        hide_zero()				#����0
        set_tab_color()			#���ñ�����ɫ
        autofilter()			#�Զ�ɸѡ
        filter_column()
        filter_column_list()

  Cell notation ��Ԫ����ʾ��(����-����) 
  	
	Excel::Writer::XLSX֧��������ʽ�ı�ʾ����ָ����Ԫ����λ��:��-�б�ʾ����A1��ʾ����

    Row-column notation uses a zero based index for both row and column
    while A1 notation uses the standard Excel alphanumeric sequence of
    column letter and 1-based row. ���磬:

	���б�ʾ������-�ж�ʹ����0Ϊ����������
	��A1��ʾ��ʹ�ñ�׼��Excel��ĸ��������Ϊ�У���1Ϊ������Ϊ�С����磺
	
        (0, 0)      # ��������ĵ�Ԫ����ʹ����-�б�ʾ����
        ('A1')      # The top left cell in A1 notation.


        (1999, 29)  # ��-�б�ʾ��.
        ('AD2000')  # ʹ��A1��ʾ����ͬһ��Ԫ��
#       ��Ԫ���еķ�Χ��Excel2003����A..IV
  
	�������ἰ��Ԫ�����̣���-�б�ʾ�������ã�

        for my $i ( 0 .. 9 ) {
            $worksheet->write( $i, 0, 'Hello' );    # Cells A1 to A10
        }

   
    A1��ʾ�������ֶ����ù�������ʹ�ù�ʽ�������а�����
        $worksheet->write( 'H1', 200 );
        $worksheet->write( 'H2', '=H1+1' ); #ʹ�ù�ʽ
Ecxel���磺
 ABCDEFGHIJKLMN
1
2
3
4
5
6
     �ڹ�ʽ�Ϳ��õķ�������Ҳ����ʹ��"A:A"���б�ʾ����
	
        $worksheet->write( 'A1', '=SUM(B:B)' );

  	�������׼��е�Excel::Writer::XLSL::Utility ģ�麬��A1��ʾ���İ������������磺
	
        use Excel::Writer::XLSX::Utility;

        ( $row, $col ) = xl_cell_to_rowcol( 'C2' );    # (1, 2)
        $str           = xl_rowcol_to_cell( 1, 2 );    # C2

  	�򵥵أ��������������½��й������������õĲ����б�������-�б�ʾ�����κ������£�������ʹ��A1��ʾ��
	
  
	ע�⣺��Excel��Ҳ����ʹ��R1C1��ʾ������Excel::Writer::XLSX��֧���⡣
	
  write( $row, $column, $token, $format )
 
	Excel����������֮�������𣬱����ַ��������֣��ո񣬹�ʽ�ͳ����ӡ�Ϊ�˼���д�����ݵĴ�����write()����Ϊ�������ض�����ָ��һ���ձ��ı�����

        write_string()
        write_number()
        write_blank()
        write_formula()
        write_url()
        write_row()
        write_col()

	һ���������ǣ��������ݿ�������ʲô�Ǿ�д��ʲô���������� ��-�б�ʾ����A1��ʾ��д�����ӣ�
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

        #���������˱���ǰ��0���ԣ�
		
        $worksheet->write( 'A16', 2                      ); # write_number()
        $worksheet->write( 'A17', 02                     ); # write_string()
        $worksheet->write( 'A18', 00002                  ); # write_string()

        # Write an array formula. Not available in Spreadsheet::WriteExcel.
         $worksheet->write( 'A19', '{=SUM(A1:B1*A2:B2)}'  ); # write_formula()

   
	"��������"�Ĺ�������������ʽ���壺
    "write_number()" ���� $token ��һ�������������������֣�
    "$token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/".

    "write_string()" ����������"����ǰ��0"("keep_leading_zeros()")���� $token ��һ���������������Ĵ���ǰ��0��������"$token =~/^0\d+$/".

    "write_blank()" ���� $token δ��������һ�����ַ���: "undef", "" �� ''.

    "write_url()" ���� $token ��һ���������������� http, https, ftp ���� mailto URL 
	              "$token =~ m|^[fh]tt?ps?://|" or "$token =~m|^mailto:|".

    "write_url()" ���� $token ��һ�����������������ڲ��Ļ��ⲿ�����ã�
                	"$token =~ m[^(in|ex)ternal:]".

    "write_formula()" ����$token�ĵ�һ���ַ���"=".

    "write_array_formula()" ���� $token ƥ�� "/^{=.*}$/".

    "write_row()" ���� $token ��һ����������.

    "write_col()" ���� $token �����������е��������á�

    "write_string()" ����ǰ����һ������������.

     $format �����ǿ�ѡ�ġ���Ӧ���Ǹ��Ϸ��ĸ�ʽ���󡣲鿴 "CELL FORMATTING":

        my $format = $workbook->add_format();
        $format->set_bold();
        $format->set_color( 'red' );
        $format->set_align( 'center' );

        $worksheet->write( 4, 0, 'Hello', $format );    # Formatted string

    write()���������Կ��ַ�����"undef"�������ṩ�˸�ʽ�����������ۣ��㲻�ص��ĶԿ�ֵ��δ����ֵ�Ĵ������鿴"write_blank()" ������
     "write()"������һ�������ǣ�ż�������ݿ�������һ���������㲻����������һ�����֡����磬����������ID�ų���ǰ��0��ͷ����������������Ϊ����д�룬��ǰ��0�ᱻɾ����������ʹ��"keep_leading_zeros()"�����ı���Ĭ����Ϊ������������������ʱ���κδ���ǰ��0�������ᱻ�����ַ�������ǰ��0���������鿴"keep_leading_zeros()"�½ڻ�ȡ����������ϸ��Ϣ��
    
	��Ҳ����ʹ��"add_write_handler()"�����Լ������ݴ��������ӵ�"write()"������

    "write()"����Ҳ�ᴦ��UTF-8��ʽ��Unicode�ַ�����
    "write" ��������:

        0 �ɹ�.
       -1 ������������
       -2 �л��г���
       -3 �ַ�������

  write_number( $row, $column, $number, $format )
    ���к���ָ����$row and $column���ĵ�Ԫ����д�������򸡵�����
        $worksheet->write_number( 0, 0, 123456 );
        $worksheet->write_number( 'A2', 2.3451 );

     $format ������ѡ.

    һ���أ�ʹ��"write()"�������㹻�ˡ�
    ע�⣺��Щ�汾��Excel2007����ʾ��Excel::Writer::XLSXд���Ĺ�ʽ����ֵ��������Excel�����п��÷��������޸������⡣
  write_string( $row, $column, $string, $format )
   
    ���к���ָ���ĵ�Ԫ����д���ַ�����
        $worksheet->write_string( 0, 0, 'Your text here' );
        $worksheet->write_string( 'A2', 'or here' );

   �������ַ�������Ϊ32767���ַ���Ȼ����Excel��Ԫ��������ʾ�������ֶ���1000�������е�32767���ַ�������ʾ��һ����ʽ���С�
    $format�����ǿ�ѡ��.

   
    "write()" ����Ҳ�ᴦ��UTF-8��ʽ���ַ��������鿴"unicode_*.pl"����
   
    һ���أ�ʹ��"write()" �������㹻�ˡ�	Ȼ��������ʱ�����ܻ�ʹ��"write_string()"����ȥд�뿴���������ֵ����ֲ��������������ֵ����ݡ����磬�����������绰���룺
        # ��Ϊ��ͨ���ַ���д��
        $worksheet->write_string( 'A1', '01209' );

   Ȼ���������û��༭���ַ�����Excel���ܻ����ַ���ת�������֡�������ʹ��Excel���ı���ʽ"@"��������:
        # ��ʽ��Ϊ�ַ���.�༭ʱ��ת��Ϊ���֡�
        my $format1 = $workbook->add_format( num_format => '@' );
        $worksheet->write_string( 'A2', '01209', $format1 );

    write_rich_string( $row, $column, $format, $string, ..., $cell_format )

    "write_rich_string()"��������д�����ж��ָ�ʽ���ַ��������磬д���ַ���"This is bold and this is italic" ������ʹ�������ķ����� 
        my $bold   = $workbook->add_format( bold   => 1 );
        my $italic = $workbook->add_format( italic => 1 );

        $worksheet->write_rich_string( 'A1',
            'This is ', $bold, 'bold', ' and this is ', $italic, 'italic' );

    
    ���������ǰ��ַ����ֶβ���$format��ʽ��������������ʽ����Ƭ��֮ǰ�����磺
        # δ��ʽ�����ַ���
          'This is an example string'

        # �ָ�
          'This is an ', 'example', ' string'

        # ��������ʽ����Ƭ��ǰ���Ӹ�ʽ
          'This is an ', $format, 'example', ' string'

        # In Excel::Writer::XLSX.
        $worksheet->write_rich_string( 'A1',
            'This is an ', $format, 'example', ' string' );

     û�и�ʽ���ַ���Ƭ��ʹ��Ĭ�ϵĸ�ʽ�����磬��д���ַ���"Some bold text"����ʹ�������ĵ�һ�����ӣ����������ڶ������ӵȼۡ�
        # ʹ��Ĭ�ϸ�ʽ:
        my $bold    = $workbook->add_format( bold => 1 );

        $worksheet->write_rich_string( 'A1',
            'Some ', $bold, 'bold', ' text' );

        # ������ȷ��:
        my $bold    = $workbook->add_format( bold => 1 );
        my $default = $workbook->add_format();

        $worksheet->write_rich_string( 'A1',
            $default, 'Some ', $bold, 'bold', $default, ' text' );

    ����Excel��ֻ�и�ʽ�������������������������񣬴�С���»��ߣ���ɫ��Ч����Ӧ�õ��ַ���Ƭ���ϡ��������������߿򣬱��������뷽ʽ���뱻Ӧ���ڵ�Ԫ����
	
    "write_rich_string()"����������������һ��������Ϊ��Ԫ����ʽʹ�ã���������һ����ʽ�����Ļ������������Ϲ��ܡ�������������ʹ��Ԫ���е�rich string ���ж��롣

        my $bold   = $workbook->add_format( bold  => 1 );
        my $center = $workbook->add_format( align => 'center' );

        $worksheet->write_rich_string( 'A5',
            'Some ', $bold, 'bold text', ' centered', $center );

    �鿴"rich_strings.pl" ��ȡ��ϸ��Ϣ

        my $bold   = $workbook->add_format( bold        => 1 );
        my $italic = $workbook->add_format( italic      => 1 );
        my $red    = $workbook->add_format( color       => 'red' );
        my $blue   = $workbook->add_format( color       => 'blue' );
        my $center = $workbook->add_format( align       => 'center' );
        my $super  = $workbook->add_format( font_script => 1 );


        # ʹ�ö��ָ�ʽд��һЩ�ַ���
        $worksheet->write_rich_string( 'A1',
            'This is ', $bold, 'bold', ' and this is ', $italic, 'italic' );

        $worksheet->write_rich_string( 'A3',
            'This is ', $red, 'red', ' and this is ', $blue, 'blue' );

        $worksheet->write_rich_string( 'A5',
            'Some ', $bold, 'bold text', ' centered', $center );

        $worksheet->write_rich_string( 'A7',
            $italic, 'j = k', $super, '(n-1)', $center );

    ���� "write_sting()" һ��������д���������ַ����� 32767��. 

  keep_leading_zeros()

    ��ʹ��"write()"����ʱ�� keep_leading_zeros()�����ı��˴���ǰ��0������Ĭ�ϴ�����ʽ��
     "write()"����ʹ����������ʽ������д��ʲô�������ݵ�Excel�������С��������ݿ���������������ʹ��"write_number()"����д�����֡��÷�����һ��������ż�����ݿ����������ֵ��㲻�뽫������һ�����֡�
	
   	��������������ID�ţ�����ǰ��0��ͷ�������������������ݵ�������д�룬��ǰ��0��ɾ���������ֶ���Excel����������ʱ����Ҳ��Ĭ����Ϊ��

     Ϊ�˱��������⣬������ʹ����ѡ��֮һ��д��һ����ʽ���������֡������ֵ����ַ���д����ʹ��"keep_leading_zeros()"�������ı�"write()"������Ĭ����Ϊ��
        # ��ʽ��д��һ������,ǰ��0��ɾ��: 1209
        $worksheet->write( 'A1', '01209' );

        #ʹ�ø�ʽд����0����������: 01209
        my $format1 = $workbook->add_format( num_format => '00000' );
        $worksheet->write( 'A2', '01209', $format1 );

        # ��ʽ�ص����ַ���д��: 01209
        $worksheet->write_string( 'A3', '01209' );

        # ��ʽ�ص����ַ���д��: 01209
        $worksheet->keep_leading_zeros();
        $worksheet->write( 'A4', '01209' );

  
	�����Ĵ���������һ��������ʾ�Ĺ�����:

         -----------------------------------------------------------
        |   |     A     |     B     |     C     |     D     | ...
         -----------------------------------------------------------
        | 1 |      1209 |           |           |           | ...
        | 2 |     01209 |           |           |           | ...
        | 3 | 01209     |           |           |           | ...
        | 4 | 01209     |           |           |           | ...

   
    �����ﵥԪ���ڲ�ͬ�ı�������ΪExcelĬ���������뷽ʽ��ʽ�ַ��������Ҷ��뷽ʽ��ʽ���֡�

    Ӧ��ע�����������û��༭�����е�"A3"��"A4"���ݣ��ַ������ָ�Ϊ���֡��⻹��Excel��Ĭ����Ϊ��ʹ���ı���ʽ"@"���Ա�������Ϊ��
        # Format as a string (01209)
        my $format2 = $workbook->add_format( num_format => '@' );
        $worksheet->write_string( 'A5', '01209', $format2 );

   
    "keep_leading_zeros()"����Ĭ���ǹرյģ�����0��1Ϊ����������û�и���ָ��������Ĭ��Ϊ1��
        $worksheet->keep_leading_zeros(   )     # Set on
        $worksheet->keep_leading_zeros( 1 );    # Set on
        $worksheet->keep_leading_zeros( 0 );    # Set off

   
  write_blank( $row, $column, $format )
   
    д�����к���ָ���Ŀհ׵�Ԫ����
        $worksheet->write_blank( 0, 0, $format );
    
    �÷��������򲻺��ַ���������ֵ�ĵ�Ԫ�����Ӹ�ʽ��
    Excel��"��"��Ԫ����"�հ�"��Ԫ����ͬ���յ�Ԫ�񲻰������ݣ��հ׵�Ԫ�񲻰������ݵ���ȴ���и�ʽ��Excel�洢"�հ�"��Ԫ�񵫺��ԡ��ա���Ԫ����
    
	��������������д���ĵ�Ԫ��Ϊ���Ҳ�����ʽ�����ᱻ���ԣ�
        $worksheet->write( 'A1', undef, $format );    # write_blank()
        $worksheet->write( 'A2', undef );             # Ignored

	�⿴�����ܷ�ζ����ʵ��ζ�������Բ����ر�����"undef"�Ϳ��ַ���ֵд�����������ݡ�
    

  write_row( $row, $column, $array_ref, $format )
   "write_row()"���������ڽ�1ά������2ά���������ݺ�Ϊһ�塣�����ڽ����ݿ���ѯ����ת��ΪExcel�����������á� �����봫�ݸ���������һ�����ö��������鱾����"write()" �����������е�ÿ��Ԫ�ص��ã����磺
    
        @array = ( 'awk', 'gawk', 'mawk' );
        $array_ref = \@array;

        $worksheet->write_row( 0, 0, $array_ref );

        # The above example is equivalent to:
        $worksheet->write( 0, 0, $array[0] );
        $worksheet->write( 0, 1, $array[1] );
        $worksheet->write( 0, 2, $array[2] );

    ע�⣺Ϊ�����������������ݵĲ������������ã���"write()"������"write_row()"����Ϊһ�������ˣ�������2�ַ������õȼۣ�
        $worksheet->write_row( 'A1', $array_ref );    # Write a row of data
        $worksheet->write(     'A1', $array_ref );    # Same thing

    �������е�write����һ����$format�����ǿ�ѡ��.����ָ����һ����ʽ�����ᱻӦ�õ���������������Ԫ���ϡ� 

   
    �����е��������ûᱻ�����С���������һ��д��2ά���������ݣ����磺
        @eec =  (
                    ['maggie', 'milly', 'molly', 'may'  ],
                    [13,       14,      15,      16     ],
                    ['shell',  'star',  'crab',  'stone']
                );

        $worksheet->write_row( 'A1', \@eec );

	���������µĹ�������
         -----------------------------------------------------------
        |   |    A    |    B    |    C    |    D    |    E    | ...
         -----------------------------------------------------------
        | 1 | maggie  | 13      | shell   | ...     |  ...    | ...
        | 2 | milly   | 14      | star    | ...     |  ...    | ...
        | 3 | molly   | 15      | crab    | ...     |  ...    | ...
        | 4 | may     | 16      | stone   | ...     |  ...    | ...
        | 5 | ...     | ...     | ...     | ...     |  ...    | ...
        | 6 | ...     | ...     | ...     | ...     |  ...    | ...

  
    ����-��˳��д�����ݣ����ο������� "write_col()"������
   
    ���Ƕ�����Ӧ��һ�ָ�ʽ�����������е��κ�"δ����"ֵ�������ԣ����������£�һ����ʽ�����Ŀհ׵�Ԫ���ᱻд�롣2�������£��ʵ����к��е�ֵ��Ȼ�����ӡ�
   
    
    ��д�����ݵ�Ԫ��ʱ��"write_row()" �������س��ֵĵ�һ������������û�г��ִ��󣬷���0��
    ���鿴 "write_arrays.pl" ������


    "write_row()"�������ı��ļ���Excel�ļ�֮���������¹���ת����
        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        my $workbook  = Excel::Writer::XLSX->new( 'file.xlsx' );
        my $worksheet = $workbook->add_worksheet();

        open INPUT, 'file.txt' or die "Couldn't open file: $!";

        $worksheet->write( $. -1, 0, [split] ) while <INPUT>;

  write_col( $row, $column, $array_ref, $format )
     "write_col()"����������һ���Ե�д��һά����ά���������ݡ������ڽ����ݿ���ѯ����ת��ΪExcel�����������á������봫��һ�����������ö������鱾���� "write()" �������������ݵ�ÿ��Ԫ�ص��á� ���磬
	 
    "write_col()"�����ܱ�����һ��д��1ά��2ά���������ݡ������ڽ����ݿ���ѯ����ת��ΪExcel�����������á������봫�ݸ���������һ�����ö��������鱾����"write()" ���������������е�ÿ��Ԫ�ص��ã����磺
        @array = ( 'awk', 'gawk', 'mawk' );
        $array_ref = \@array;

        $worksheet->write_col( 0, 0, $array_ref );

        # ���������ӵȼ���:
        $worksheet->write( 0, 0, $array[0] );
        $worksheet->write( 1, 0, $array[1] );
        $worksheet->write( 2, 0, $array[2] );

    �������е�write����һ����$format�����ǿ�ѡ��.����ָ���˸�ʽ�����ᱻӦ�õ��������е�Ԫ���ϡ� 


    �����е��������ûᱻ���������Դ�����������һ��д��2ά���������ݡ����磺
        @eec =  (
                    ['maggie', 'milly', 'molly', 'may'  ],
                    [13,       14,      15,      16     ],
                    ['shell',  'star',  'crab',  'stone']
                );

        $worksheet->write_col( 'A1', \@eec );

    ���������µĹ�������

         -----------------------------------------------------------
        |   |    A    |    B    |    C    |    D    |    E    | ...
         -----------------------------------------------------------
        | 1 | maggie  | milly   | molly   | may     |  ...    | ...
        | 2 | 13      | 14      | 15      | 16      |  ...    | ...
        | 3 | shell   | star    | crab    | stone   |  ...    | ...
        | 4 | ...     | ...     | ...     | ...     |  ...    | ...
        | 5 | ...     | ...     | ...     | ...     |  ...    | ...
        | 6 | ...     | ...     | ...     | ...     |  ...    | ...

 
	����������-�е�˳��д�����鿴������"write_row()"������

    ���Ƕ�����Ӧ��һ�ָ�ʽ�����������е��κ�"δ����"ֵ�������ԣ����������£�һ����ʽ�����Ŀհ׵�Ԫ���ᱻд�롣2�������£��ʵ����к��е�ֵ��Ȼ�����ӡ�


	����������˵�ģ�"write()" �����ܱ����� "write_row()" ��ͬ���ʣ����� "write_row()"������Ƕ�׵��������õ�����.
   
    ���ԣ�������2�����������ǵȼ۵ģ����ܶ�"write_col()"�������ĵ��ÿ�ά���Ը��á�
        $worksheet->write_col( 'A1', $array_ref     ); # Write a column of data
        $worksheet->write(     'A1', [ $array_ref ] ); # Same thing

   
    ��д�����ݵ�Ԫ��ʱ��"write_col()" �������س��ֵĵ�һ������������û�г��ִ��󣬷���0��
    �鿴�����ġ�write�����������ķ���ֵ�� 
    
    Ҳ���鿴 "write_arrays.pl"������

  write_date_time( $row, $col, $date_string, $format )
   
    "write_date_time()" ������������ָ�����еĵ�Ԫ����д�����ڻ�ʱ�䣺
        $worksheet->write_date_time( 'A1', '2004-05-13T23:20', $date_format );

    $date_string Ӧ�������еĸ�ʽ:

        yyyy-mm-ddThh:mm:ss.sss

    ������ISO8601���ڵ���Ӧ��ע�����ǲ�֧��ȫ����Χ�ڵ�ISO8601��ʽ��
    �����ı�����$date_string �����������ģ�
        yyyy-mm-ddThh:mm:ss.sss         # Standard format
        yyyy-mm-ddT                     # No time
                  Thh:mm:ss.sss         # No date
        yyyy-mm-ddThh:mm:ss.sssZ        # Additional Z (but not time zones)
        yyyy-mm-ddThh:mm:ss             # No fractional seconds
        yyyy-mm-ddThh:mm                # No seconds

    ע��"T"�������������Ǳ����ġ�
   
    ����Ӧ��һֱ�и�$format��ʽ������������������ʽ���֡������ǵ��͵����ӣ�
        my $date_format = $workbook->add_format( num_format => 'mm/dd/yy' );
        $worksheet->write_date_time( 'A1', '2004-05-13T23:20', $date_format );

     ����1900��Ԫ���Ϸ�������Ӧ����1900-01-01��9999-12-31֮�䣬������1904��Ԫ���Ϸ���������1904-01-01 �� 9999-12-31������Excel��������Щ��Χ�����ڻᱻ�����ַ���д�롣
    ���鿴 date_time.pl������

  write_url( $row, $col, $url, $format, $label )
     ��URL�ĳ�����д�����к���ָ���ĵ�Ԫ���С���������2�������ɣ��ɼ��ı��ǺͲ��ɼ������ӡ��ɼ��ı���������һ��������ָ���˿��滻�ı��ǡ�$label�����ǿ�ѡ�ġ�������"write()"����д�롣���ˣ����Խ��ַ��������ֻ���ʽ��Ϊ���ǡ�
    
	$format����Ҳ�ǿ�ѡ�ģ�Ȼ����û�и�ʽ�Ļ������ӾͲ���һ����ʽ�ˡ�

    �����ĸ�ʽ��:

        my $format = $workbook->add_format( color => 'blue', underline => 1 );

    ע�⣬�����û�û��ָ��һ�ָ�ʽ������Ϊ��Spreadsheet::WriteExcel�ṩ��Ĭ�ϵĳ����Ӹ�ʽ��ͬ��

	֧��4��web������URI "http://", "https://", "ftp://" �� "mailto:":

        $worksheet->write_url( 0, 0, 'ftp://www.perl.org/', $format );
        $worksheet->write_url( 1, 0, 'http://www.perl.com/', $format, 'Perl' );
        $worksheet->write_url( 'A3', 'http://www.perl.com/',      $format );
        $worksheet->write_url( 'A4', 'mailto:jmcnamara@cpan.org', $format );

    There are two local URIs supported: "internal:" and "external:". These
    are used for hyperlinks to internal worksheet references or external
    workbook and worksheet references:
    ֧��2�ֱ���URLs��"internal:" �� "external:����Щ�����ڲ����������õĳ����ӻ��ⲿ�������������������ã�
        $worksheet->write_url( 'A6',  'internal:Sheet2!A1',              $format );
        $worksheet->write_url( 'A7',  'internal:Sheet2!A1',              $format );
        $worksheet->write_url( 'A8',  'internal:Sheet2!A1:B2',           $format );
        $worksheet->write_url( 'A9',  q{internal:'Sales Data'!A1},       $format );
        $worksheet->write_url( 'A10', 'external:c:\temp\foo.xlsx',       $format );
        $worksheet->write_url( 'A11', 'external:c:\foo.xlsx#Sheet2!A1',  $format );
        $worksheet->write_url( 'A12', 'external:..\foo.xlsx',            $format );
        $worksheet->write_url( 'A13', 'external:..\foo.xlsx#Sheet2!A1',  $format );
        $worksheet->write_url( 'A13', 'external:\\\\NET\share\foo.xlsx', $format );

  
	���͵Ĺ�����������ʽ��"Sheet1!A1"��������ʹ�ñ�׼��Excel��ʾ��"Sheet1!A1:B2"��ָ����������Χ��

    In external links the workbook and worksheet name must be separated by
    the "#" character: "external:Workbook.xlsx#Sheet1!A1'".
	�ⲿ�����У��������͹����������ֱ�����"#"������"external:Workbook.xlsx#Sheet1!A1'"

    You can also link to a named range in the target worksheet. ���磬
    say you have a named range called "my_name" in the workbook
    "c:\temp\foo.xlsx" you could link to it as follows:
    ��Ҳ�������ӵ�Ŀ�깤�����е�һ������ֵ���ϡ����磬������"c:\temp\foo.xlsx"������������һ��������ֵ������"my_name"�������԰������ķ������ӵ�����
        $worksheet->write_url( 'A14', 'external:c:\temp\foo.xlsx#my_name' );

    ExcelҪ�������ո�������ĸ�ַ��Ĺ�������Ҫ�õ�����������������"'Sales Data'!A1"��������ʹ�õ������������ַ�������������������Ҫʹ��\'ת�嵥���ţ���ʹ��q{}��
  
    Ҳ֧�ֵ������ļ������ӡ�MS/Novell �����ļ�һ����2����б�ܿ�ͷ������"\\NETWORK\etc"��Ϊ���ڵ����Ż�˫�����ַ�������������������������Ҫת�巴б�ܣ�'\\\\NETWORK\etc'.

    ������ʹ��˫�����ַ�������Ӧ��ע��ת���κο�������Ԫ�ַ����ַ���������Ϣ���鿴Perlfaq5��Ϊʲô������DOS·����ʹ�� "C:\temp\foo" in DOS paths?"��
  
  
    ������������ʹ����б���������󲿷��������⡣��б�����ڲ���ת���ɷ�б�ܣ�
        $worksheet->write_url( 'A14', "external:c:/temp/foo.xlsx" );
        $worksheet->write_url( 'A15', 'external://NETWORK/share/foo.xlsx' );

    

  write_formula( $row, $column, $formula, $format, $value )
	����ʽ������д�����к���ָ���ĵ�Ԫ���С�
        $worksheet->write_formula( 0, 0, '=$B$3 + B4' );
        $worksheet->write_formula( 1, 0, '=SIN(PI()/4)' );
        $worksheet->write_formula( 2, 0, '=SUM(B1:B5)' );
        $worksheet->write_formula( 'A4', '=IF(A3>1,"Yes", "No")' );
        $worksheet->write_formula( 'A5', '=AVERAGE(1, 2, 3, 4)' );
        $worksheet->write_formula( 'A6', '=DATEVALUE("1-Jan-2001")' );

    ͬ��Ҳ֧�����鹫ʽ:

        $worksheet->write_formula( 'A7', '{=SUM(A1:B1*A2:B2)}' );

 
    ������Ҫ������ָ����ʽ�ļ������������벻�ܼ��㹫ʽֵ�ķ�ExcelӦ�ó���һ������ʱ����ż����Ҫ����������$valueֵ�������ڲ����б���ĩβ��
        $worksheet->write( 'A1', '=2+2', $format, 4 );


  write_array_formula($first_row, $first_col, $last_row, $last_col, $formula, $format, $value)
     �����鹫ʽд�뵽һ����Ԫ��ֵ���С���Excel��һ�����鹫ʽ������һ��ֵ��ִ�м����Ĺ�ʽ�������Է��ص���ֵ��һ��ֵ����
    ��ʽ���ߵ�һ�Ի����ű�������һ�����鹫ʽ��"{=SUM(A1:B1*A2:B2)}"���������鹫ʽ���ص���ֵ����$first_ �� $last_ ����Ӧ��һ����
        $worksheet->write_array_formula('A1:A1', '{=SUM(B1:C1*B2:C2)}');


	���������½���ʹ��"write_formula()"�� "write()"����������Щ��

        # ��������һ�������Ǹ����ࣺ
        $worksheet->write( 'A1', '{=SUM(B1:C1*B2:C2)}' );
        $worksheet->write_formula( 'A1', '{=SUM(B1:C1*B2:C2)}' );

    For array formulas that return a range of values you must specify the
    range that the return values will be written to:

        $worksheet->write_array_formula( 'A1:A3',    '{=TREND(C1:C3,B1:B3)}' );
        $worksheet->write_array_formula( 0, 0, 2, 0, '{=TREND(C1:C3,B1:B3)}' );

   ������Ҫ������ָ����ʽ�ļ������������벻�ܼ��㹫ʽֵ�ķ�ExcelӦ�ó���һ������ʱ����ż����Ҫ����������$valueֵ�������ڲ����б���ĩβ��

        $worksheet->write_array_formula( 'A1:A3', '{=TREND(C1:C3,B1:B3)}', $format, 105 );

     ���⣬һЩExcel2007�����ڰ汾û���ṩ���鹫ʽ���������ǲ��������鹫ʽ��ֵ����װ���µ�Office Service Pack�����޸������⡣
    �鿴 "array_formula.pl" ������

   ע�⣺Spreadsheet::WriteExcel��֧�����鹫ʽ��
  store_formula( $formula )
   
    ����ʹ�á�����һ��Spreadsheet::WriteExcel ������Excel::Writer::XLSX�Ѿ�������Ҫ�ˡ������档
  repeat_formula( $row, $col, $formula, $format )
 
     ����ʹ�á�����һ��Spreadsheet::WriteExcel ������Excel::Writer::XLSX.�Ѿ�������Ҫ�ˡ�

    �� Spreadsheet::WriteExcel��д�빫ʽ�������ǰ����ģ���Ϊ�����ɵݹ��½��Ľ�����������
	"store_formula()" �� "repeat_formula()" ����������Ԥ������ʽ��Ϊ�����ظ���ʽ��ϵͳ�����ķ���֮һ��
   
    ��Excel::Writer::XLSX�У��ⲻ����Ҫ����Ϊд�빫ʽ����д���ַ���������һ���졣
    The methods remain for backward compatibility but new
    Excel::Writer::XLSX programs shouldn't use them.
    ���Ϸ����������ݣ������µ�Excel::Writer::XLSX ����ʹ�����ǡ�
  write_comment( $row, $column, $string, ... )
  
    "write_comment()" ������������Ԫ����д��ע�͡�
	��Ԫ���е�ע����Excel��Ԫ�����Ͻǵ�С��ɫ���Ǳ�ע���������Ƶ���ɫ�����Ͻ���ʾע�͡� 
    ������������ʽ�������ڵ�Ԫ��������ע�ͣ�
        $worksheet->write        ( 2, 2, 'Hello' );
        $worksheet->write_comment( 2, 2, 'This is a comment.' );


	ͨ������������һ��"A1"��Ԫ�����ô��� $row �� $column ������

        $worksheet->write        ( 'C3', 'Hello');
        $worksheet->write_comment( 'C3', 'This is a comment.' );

     "write_comment()" ����Ҳ����UTF-8��ʽ���ַ�����

        $worksheet->write_comment( 'C3', "\x{263a}" );       # Smiley
        $worksheet->write_comment( 'C4', 'Comment ca va?' );

 
    ���˻�����3��������ʽ��"write_comment()"�������Դ��ݼ��Կ�ѡ�� ��ֵ��������ע�͵ĸ�ʽ��
        $worksheet->write_comment( 'C3', 'Hello', visible => 1, author => 'Perl' );

   
    ��Щѡ���Ĵ��������ر𣬲���ͨ��Ĭ�ϵ�ע����Ϊ��������Ҫ�ġ�Ȼ������������Ҫ�Ե�Ԫ��ע�͸��õĿ��ƣ���������ѡ������ʹ�ã�
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
    
        ��ѡ�����ڱ���˭�Ǹõ�Ԫ��ע�͵����ߡ�Excel�ڹ������ײ���״̬��������ʽע�͵����ߡ�
            $worksheet->write_comment( 'C3', 'Atonement', author => 'Ian McEwan' );


        ���еĵ�Ԫ��ע��Ĭ�ϵ�������ʹ�� "set_comments_author()"�������ã������棺
            $worksheet->set_comments_author( 'Perl' );

    Option: visible
    
        ��ѡ�������ڵ��򿪹�����ʱ��ʹ��Ԫ��ע�Ϳɼ���Excel��Ĭ�ϵ���Ϊ��ע�ͱ����ء�Ȼ������Excel��Ҳ����ʹ����ע�ͻ�����ע�Ϳɼ�����Excel::Writer::XLSX�У����԰������ķ���ʹ����ע�Ϳɼ���
            $worksheet->write_comment( 'C3', 'Hello', visible => 1 );

      
        ʹ��"show_comments()" �����������������棩��ʹ�������е�����ע�Ϳɼ���
		��Ӧ�أ��������еĵ�Ԫ��ע�Ͷ��ɼ��ˣ����������ص���ע�ͣ�
            $worksheet->write_comment( 'C3', 'Hello', visible => 0 );

    Option: x_scale
        This option is used to set the width of the cell comment box as a
        factor of the default width.
        ��ѡ���������õ�Ԫ��ע�Ϳ��Ŀ��ȣ�
            $worksheet->write_comment( 'C3', 'Hello', x_scale => 2 );
            $worksheet->write_comment( 'C4', 'Hello', x_scale => 4.2 );

    Option: width
     
        ��ѡ���������õ�Ԫ��ע�Ϳ��Ŀ��ȣ������ر�ʾ
            $worksheet->write_comment( 'C3', 'Hello', width => 200 );

    Option: y_scale
        This option is used to set the height of the cell comment box as a
        factor of the default height.
        ��ѡ���������õ�Ԫ��ע�Ϳ��ĸ߶ȣ�
            $worksheet->write_comment( 'C3', 'Hello', y_scale => 2 );
            $worksheet->write_comment( 'C4', 'Hello', y_scale => 4.2 );

    Option: height
        This option is used to set the height of the cell comment box
        explicitly in pixels.
         ��ѡ���������õ�Ԫ��ע�Ϳ��ĸ߶ȣ������ر�ʾ
            $worksheet->write_comment( 'C3', 'Hello', height => 200 );

    Option: color

        ��ѡ���������õ�Ԫ��ע�Ϳ��ı���ɫ��������ʹ��Excel::Writer::XLSX ��ʶ���ľ�����ɫ����ɫ������
            $worksheet->write_comment( 'C3', 'Hello', color => 'green' );
            $worksheet->write_comment( 'C4', 'Hello', color => 0x35 );      # Orange

    Option: start_cell
   
        ��ѡ����������ע�ͽ��������ĸ���Ԫ����By default Excel displays comments one cell to the right and
        one cell above the cell to which the comment relates.
		Ȼ���������Ըı�����Ĭ����Ϊ��������Ը���Ļ����������������У�Ĭ�ϻ��ڵ�Ԫ����D2���г��ֵ�ע�ͻ��ƶ�����E2���С�
            $worksheet->write_comment( 'C3', 'Hello', start_cell => 'E2' );
2011-12-5
    Option: start_row
    
        ��ѡ����������ע�ͽ�����������һ�С��д�0��ʼ������
            $worksheet->write_comment( 'C3', 'Hello', start_row => 0 );

    Option: start_col
   
		��ѡ����������ע�ͽ�������һ�г��֡��д�0��ʼ������

            $worksheet->write_comment( 'C3', 'Hello', start_col => 4 );

    Option: x_offset
        This option is used to change the x offset, in pixels, of a comment
        within a cell:
        	��ѡ�����ڸı䵥Ԫ����ע�͵�x�᷽����ƫ�����������ؼ��㡣
            $worksheet->write_comment( 'C3', $comment, x_offset => 30 );

    Option: y_offset
        This option is used to change the y offset, in pixels, of a comment
        within a cell:
		��ѡ�����ڸı䵥Ԫ����ע�͵�y�᷽����ƫ�����������ؼ��㡣

            $worksheet->write_comment('C3', $comment, x_offset => 30);

    

	ע�⣬ʹ������start_cell, start_row, start_col, x_offset �� y_offset��ѡ��������Ԫ��ע��λ�ã�Excel�����ڵ�Ԫ���ɼ�ʱ����ʾ��Ԫ��ע�͵�ƫ�����������������Ƶ���������ʱ��Excel����ʾ���صĵ�Ԫ����

   
	ע���иߺ�ע�͡�������ָ������ע�͵ĵ�Ԫ�����иߣ���Excel::Writer::XLSX������ע�͵ĸ߶��Ա���Ĭ�ϸ߶Ȼ��û�ָ���ĳߴ硣Ȼ���������������ı������Ի��ڵ�Ԫ����ʹ���˺ܴ������壬Excel���Զ������иߡ�����ζ���и߶�������ʱ��ģ����δ֪�ģ�����ע�Ϳ�������һ����չ��ʹ�� "set_row()"������ʽ��ָ���и��������������⡣

  show_comments()

    �÷������ڵ��򿪹�����ʱ�������е�Ԫ��ע�Ϳɼ�
        $worksheet->show_comments();


    ʹ��"write_comment"������"visible"������ʹ������ע�Ϳɼ��������棩��
        $worksheet->write_comment( 'C3', 'Hello', visible => 1 );

	�������еĵ�Ԫ��ע�Ͷ��ɼ��ˣ������԰������ķ������ص���ע�ͣ�

        $worksheet->show_comments();
        $worksheet->write_comment( 'C3', 'Hello', visible => 0 );

  set_comments_author()
    �÷����������õ�Ԫ��ע�͵�Ĭ�����ߡ�
        $worksheet->set_comments_author( 'Perl' );

    ʹ��"write_comment"������"author"�������õ���ע�͵����ߣ������棩��

	����û��ָ�����ߣ�Ĭ�ϵ�ע�������ǿ��ַ���,''.

  add_write_handler( $re, $code_ref )
   
	�÷���������չ Excel::Writer::XLSX ��"write()"�����������û����������ݡ�

	�������鿴�����½ڵ�"write()"���������ᷢ�����Ǽ�����������"write_*"�����ı�����Ȼ��������������������Ը����ȷ��

 
    һ�ַ��������Լ������������ݲ����ú��ʵ� "write_*" ������
	��һ�ַ�����ʹ��"add_write_handler()"���������Լ����Զ�����Ϊ���ӵ�"write()"�����С�
  
  
	"add_write_handler()" ��2��������$re,һ��ƥ���������ݵ���������ʽ��
	����$code_ref��һ���ص�������������ƥ���������ݣ�

        $worksheet->add_write_handler( qr/^\d\d\d\d$/, \&my_write );

    (����Щ������"qr"����������������������ʽ�ַ���).


	�÷���ʹ�����¡���������д��7�����ֵ�ID����Ϊ�ַ����������κ�ǰ��0�������԰�����������

        $worksheet->add_write_handler( qr/^\d{7}$/, \&write_my_id );


        sub write_my_id {
            my $worksheet = shift;
            return $worksheet->write_string( @_ );
        }

    * ��Ҳ����ʹ��"keep_leading_zeros()"����.

	Ȼ����������ʹ��һ�����ʵ��ַ�������"write()"���������ᱻ�Զ�������
        # д�� 0000000.�����أ����ᱻд������0��
        $worksheet->write( 'A1', '0000000' );

	�ص�����������һ���Ա����ù����������ã������������еĲ��������ݸ�"write()".�ص������ῴ��������ʾ��@_�����б���

        $_[0]   A ref to the calling worksheet. *
        $_[1]   Zero based row number.
        $_[2]   Zero based column number.
        $_[3]   A number or string or token.
        $_[4]   A format ref if any.
        $_[5]   Any other arguments.
        ...

        *  It is good style to shift this off the list so the @_ is the same
           as the argument list seen by write().


	���Ļص�����Ӧ��ʹ��"return()"����"write_*" �����ķ���ֵ��
	�򷵻ء�undef�����������ܾ���ƥ�䲢������"write()"��������������

    So ���磬 if you wished to apply the previous filter only to ID
    values that occur in the first column you could modify your callback
    function as follows:
	���ԣ����磬����������ǰ���Ĺ���ֻӦ�õ������ڵ�һ�е�IDֵ�ϣ������԰��������޸����Ļص�������

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
    ���ڣ������ڡ���һ�к��������ϵõ���ͬ����Ϊ��
        $worksheet->write( 'A1', '0000000' );    # Writes 0000000
        $worksheet->write( 'B1', '0000000' );    # Writes 0

	���������Ӷ����������򣬴�ʱ���ǻᰴ�ձ����ӵ�˳�����á�

    
	ע�⣬"add_write_handler()"�ر��ʺϴ������ݡ�

    �鿴 "write_handler 1-4" ������

  insert_image( $row, $col, $filename, $x, $y, $scale_x, $scale_y )
    
	����֧�֡�Ŀǰֻ��96���ص�ͼ����Ч���������´η������޸���

   
    �÷����������������в���ͼ����ͼ����ʽ������PNG, JPEG �� BMP�� $x, $y, $scale_x �� $scale_y �����ǿ�ѡ�ġ�
    	$worksheet1->insert_image( 'A1', 'perl.bmp' );
        $worksheet2->insert_image( 'A1', '../images/perl.bmp' );
        $worksheet3->insert_image( 'A1', '.c:\images\perl.bmp' );

    
	$x �� $y������������ָ������$row��$colָ���ĵ�Ԫ�������Ͻǵ�ƫ������ƫ����ֵ�����ؼ��㣺

        $worksheet1->insert_image('A1', 'perl.bmp', 32, 10);

   	ƫ�������Ա�ͼ�������ĵ�Ԫ���ĸ߶Ȼ����ȴ�������������ͬһ����Ԫ���ж���һ��������ͼ������ż�������á�

   
	����$scale_x �� $scale_y������ˮƽ�ʹ�ֱ�ز���������ͼƬ��

        # Scale the inserted image: width x 2.0, height x 0.8
        $worksheet->insert_image( 'A1', 'perl.bmp', 0, 0, 2, 0.8 );

    �鿴��"images.pl" ������


	ע�⣺���������ı�ͼ����ռ����һ�л��е�Ĭ�ϳߴ磬��������"insert_image()"֮ǰ����"set_row()" �� "set_column()"��
	������ʹ�õ�������Ĭ�ϵĴ����еĸ߶�Ҳ���ı䡣�ⷴ����Ҳ��Ӱ����ͼ���ĳߴ硣��Ӧ��ʹ��"set_row()"��ʽ�������и��������������⣬�����������˻��ı��иߵ�������С��


	BMPͼ��Ӧ����24���أ���ɫΪ���ʣ�λͼ��ͨ��������Ӧ��ʹ��BMPͼ������Ϊ����û�б�ѹ����

  insert_chart( $row, $col, $chart, $x, $y, $scale_x, $scale_y )
 	�÷����������������в���һ��ͼ��������ͼ��������"add_chart()"�������������������������������� "embedded"ѡ�

        my $chart = $workbook->add_chart( type => 'line', embedded => 1 );

        # Configure the chart.
        ...

        # Insert the chart into the a worksheet.
        $worksheet->insert_chart( 'E2', $chart );


	�鿴"add_chart()" ��ȡ������������ͼ��������ϸ�ڣ�
	�鿴Excel::Writer::XLSX::Chart��ȡ������������ͼ����ϸ�ڡ��鿴"chart_*.pl"������

    $x, $y, $scale_x �� $scale_y �����ǿ�ѡ�� ��
	


	$x �� $y ������ָ������$row��$colָ���ĵ�Ԫ�����Ͻǵ�ƫ������ƫ����ֵ�����ؼ��㡣

        $worksheet1->insert_chart( 'E2', $chart, 3, 3 );

    
	����$scale_x �� $scale_y�����ڴ�ˮƽ�����ʹ�ֱ��������ͼ����

        # Scale the width by 120% and the height by 150%
        $worksheet->insert_chart( 'E2', $chart, 0, 0, 1.2, 1.5 );

  data_validation()
    
	"data_validation()"�������ڹ���Excel���ݼ����������û����뵽һ�������б���

        $worksheet->data_validation('B3',
            {
                validate => 'integer',   #��֤
                criteria => '>',         #��׼
                value    => 100,         #ֵ
            });

        $worksheet->data_validation('B5:B9',
            {
                validate => 'list',
                value    => ['open', 'high', 'close'],
            });

 
	�÷��������ܶ����������ڵ������½ڡ�DATA VALIDATION IN EXCEL��������ϸ������

	�鿴"data_validate.pl" ������
	
  conditional_format()
  
	"conditional_format()" ������������Ԫ���������û��Զ�����׼�ĵ�Ԫ����Χ�����Ӹ�ʽ

        $worksheet->conditional_formatting( 'A1:J10',
            {
                type     => 'cell',
                criteria => '>=',
                value    => 50,
                format   => $format1,
            }
        );


	�÷��������ܶ����������ڵ������½ڡ�CONDITIONAL FORMATTING IN EXCEL��������ϸ����

	�鿴"conditional_format.pl"������

  get_name()
   
	"get_name()" �������ڼ��������������֡����磺

        for my $sheet ( $workbook->sheets() ) {
            print $sheet->get_name();
        }

	����Excel::Writer::XLSX�����ƺ�Excel���ڲ�ԭ����û������"set_name()" ������
	���ù��������ֵ�Ψһ������ͨ��"add_worksheet()"������

  activate()

	"activate()"��������ָ����һ�����ж����������Ĺ������У���һ���������ǳ�ʼ�ɼ��ģ�

        $worksheet1 = $workbook->add_worksheet( 'To' );
        $worksheet2 = $workbook->add_worksheet( 'the' );
        $worksheet3 = $workbook->add_worksheet( 'wind' );

        $worksheet3->activate();

   
	����Excel VBA��active�������ơ�����ͨ��"select()"����ѡȡ���Ź�������
	�����棬Ȼ����ֻ��һ�Ź������Ǽ����ġ�
	
	��һ�Ź�����Ĭ���Ǽ����ġ�

  select()

	"select()"�������ڱ����Ӻ��ж��Ź������Ĺ�������ѡȡһ�ţ�

        $worksheet1->activate();
        $worksheet2->select();
        $worksheet3->select();

	��ѡ�еĹ������ı�ǩ�Ǹ����ġ�ѡȡ���Ź������ǰ�����������һ����һ�ַ��������ԣ����磬����һ�ٴ�ӡ���Ź�������ͨ��"activate()"�����������Ĺ�����Ҳ�ᱻѡ�С�

  hide()
	 "hide()" ������������һ����������

        $worksheet2->hide();

  	Ϊ�˱���ʹ���м����ݻ����������Ի��û�����������Ҫ����һ��������


	һ�����صĹ��������ܱ���������ѡ�У����Ը÷�����"activate()" �� "select()"�ǻ����ų��ġ����⣬��Ϊ��һ�Ź�����Ĭ���Ǳ�ѡ�еģ��㲻�ܲ����������Ĺ����������ص�һ�Ź�������

        $worksheet2->activate();
        $worksheet1->hide();

  set_first_sheet()
 	"activate()"������������ѡ����һ�Ź�������Ȼ���������кܶ��Ź���������ѡ�еĹ��������ܲ�����������Ļ�ϡ�������ʹ��"set_first_sheet()"����ѡ�������˿ɼ��Ĺ����������⣺

        for ( 1 .. 20 ) {
            $workbook->add_worksheet;
        }

        $worksheet21 = $workbook->add_worksheet();
        $worksheet22 = $workbook->add_worksheet();

        $worksheet21->set_first_sheet();
        $worksheet22->activate();


	�÷������Ǿ�����Ҫ��Ĭ��ֵ�ǵ�һ�Ź�������

  protect( $password, \%options )
	"protect()" �������ڷ�ֹ���������޸ģ�

        $worksheet->protect();


	"protect()"����Ҳ���Կ�����Ԫ����"locked"��"hidden"������Ӱ�죬���������˵�Ԫ����"locked"��"hidden"���ԵĻ���һ��*locked*�ĵ�Ԫ�����ܹ����༭�����Ҹ�����Ĭ�϶����е�Ԫ���ǿ����ġ�һ�����صĵ�Ԫ������ʾ��ʽ�Ľ��������ǹ�ʽ������

    
	�鿴"protection.pl" ������"CELL FORMATTING" ������"set_locked" �� "set_hidden"��ʽ������

	������ѡ���Ե�����һ�����뵽�������У�

        $worksheet->protect( 'drowssap' );

	����һ�����ַ���''�Ϳ���û�������ı���һ����

 	ע�⣺Excel�й����������������ṩ�ı����ܴ�������û�м����������ݲ��Һ����ױ�������"Excel::Writer::XLSX"��֧����ȫ�Ĺ��������ܣ���Ϊ����Ҫһ����ȫ��ͬ���ļ���ʽ������Ҫ���Ѽ��������µ�ʱ������ʵ�֡�

    
	�����Դ���һ�������κ�һ����ȫ��������ʾ��ֵ��ɢ��������ָ�����뱣���ĸ�������Ԫ�أ�

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


	������ʾ����Ĭ�ϵĲ���ֵ������Ԫ�صı�������ʹ�������ķ�����

        $worksheet->protect( 'drowssap', { insert_rows => 1 } );

  set_selection( $first_row, $first_col, $last_row, $last_col )
    This method can be used to specify which cell or cells are selected in a
    worksheet. The most common requirement is to select a single cell, in
    which case $last_row and $last_col can be omitted. The active cell
    within a selected range is determined by the order in which $first and
    $last are specified. It is also possible to specify a cell or a range
    using A1 notation. 
	�÷�������ָ����һ�Ź�������ѡ���ĸ�����Щ��Ԫ���������������ѡ��һ����Ԫ�񣬴�ʱ$last_row �� $last_col ����ʡ�ԡ���ѡ���ڵļ��Ԫ����ָ����$first and $last ��˳��������Ҳ����ʹ��A1��ʾ��ָ����Ԫ����һ����Χ��

    Examples:

        $worksheet1->set_selection( 3, 3 );          # 1. Cell D4.
        $worksheet2->set_selection( 3, 3, 6, 6 );    # 2. Cells D4 to G7.
        $worksheet3->set_selection( 6, 6, 3, 3 );    # 3. Cells G7 to D4.
        $worksheet4->set_selection( 'D4' );          # Same as 1.
        $worksheet5->set_selection( 'D4:G7' );       # Same as 2.
        $worksheet6->set_selection( 'G7:D4' );       # Same as 3.

	Ĭ�ϵĵ�Ԫ��ѡ����(0,0),'A1'.

  set_row( $row, $height, $format, $hidden, $level, $collapsed )
 
	�÷������ڸı��е�Ĭ�����ԡ���$row֮�⣬�����������ǿ�ѡ�ġ�

	�÷�������ͨ���÷��Ǹ����иߣ�

        $worksheet->set_row( 0, 20 );    # ��һ�е��и߸�Ϊ20

	�����������ø�ʽʱ�������иߣ������Դ���"undef"��Ϊ�и߲�����

        $worksheet->set_row( 0, undef, $format );

	$format�������Ա�Ӧ�õ������κ�û�и�ʽ�ĵ�Ԫ���ϣ����磺

        $worksheet->set_row( 0, undef, $format1 );    # Set the format for row 1
        $worksheet->write( 'A1', 'Hello' );           # Defaults to $format1
        $worksheet->write( 'B1', 'Hello', $format2 ); # Keeps $format2

	�������������ַ�������һ���и�ʽ����Ӧ�����κε���"write()"��������Ϊ֮ǰ���ø÷��������ø÷����Ժ��Ḳ����ǰָ�����κθ�ʽ��

   
	�����������У�$hidden����Ӧ������Ϊ1���ⱻ���ڣ����磬�ڸ��Ӽ����������м䲽�裺

        $worksheet->set_row( 0, 20,    $format, 1 );
        $worksheet->set_row( 1, undef, undef,   1 );


	 $level�������������еķּ���ʾ��outline level����"OUTLINES AND GROUPING IN EXCEL"�½��й��ڷּ���ʾ����������ͬ���ּ���ʾ���лᱻ���ϵ�һ����Ϊһ����һ�ķּ���ʾ��


	����������Ϊ��1����2����0��ʼ�����������˷ּ���ʾ1��

        $worksheet->set_row( 1, undef, undef, 0, 1 );
        $worksheet->set_row( 2, undef, undef, 0, 1 );

	����$level����һͬʹ�õ�ʱ����$hidden����Ҳ�����������۵��ķּ���ʾ�У�

        $worksheet->set_row( 1, undef, undef, 1, 1 );
        $worksheet->set_row( 2, undef, undef, 1, 1 );


	�����۵��ķּ���ʾ����Ӧ��ʹ�ÿ�ѡ��$collapsed����ָ����һ�к����۵�����"+".

        $worksheet->set_row( 3, undef, undef, 0, 0, 1 );

    �鿴 "outline.pl" �� "outline_collapsed.pl" ��


	Excel��������7���ּ���ʾ�����ˣ�$level����Ӧ���ڷ�Χ��"0 <= $level <= 7"�ڡ�

  set_column( $first_col, $last_col, $width, $format, $hidden, $level, $collapsed )
  
	�÷������ڸı䵥һ��һ����Χ���е�Ĭ�����ԡ����� $first_col �� $last_col �⣬���в������ǿ�ѡ�ġ�

 	����"set_column()"��Ӧ�õ���һ�У� $first_col�� $last_col��ֵӦ��һ����
	��$last_col��0�������£���������Ϊ��$first_colһ����ֵ��

	Ҳ�ɸ�������ʹ���е�A1��ʾ����ָ���еķ�Χ��

    ����:

        $worksheet->set_column( 0, 0, 20 );    # Column  A   width set to 20
        $worksheet->set_column( 1, 3, 30 );    # Columns B-D width set to 30
        $worksheet->set_column( 'E:E', 20 );   # Column  E   width set to 20
        $worksheet->set_column( 'F:H', 30 );   # Columns F-H width set to 30

    The width corresponds to the column width value that is specified in
    Excel. It is approximately equal to the length of a string in the
    default font of Arial 10. Unfortunately, there is no way to specify
    "AutoFit" for a column in the Excel file format. This feature is only
    available at runtime from within Excel.
	

    ͨ�� $format�����ǿ�ѡ��,������Ϣ,�鿴"CELL FORMATTING". 
	�����������ø�ʽʱ�������п��������Դ���"undef"��Ϊ�п�������

        $worksheet->set_column( 0, 0, undef, $format );

	$format�������Ա�Ӧ�õ������κ�û�и�ʽ�ĵ�Ԫ���ϣ����磺

        $worksheet->set_column( 'A:A', undef, $format1 );    # Set format for col 1
        $worksheet->write( 'A1', 'Hello' );                  # Defaults to $format1
        $worksheet->write( 'A2', 'Hello', $format2 );        # Keeps $format2

	�������������ַ�������һ���и�ʽ����Ӧ�����κε���"write()"��������Ϊ֮ǰ���ø÷������������ڵ���"write()"����֮�����ø÷��������������κ�Ч����

	Ĭ�ϵ��и�ʽ����Ĭ�ϵ��и�ʽ��

        $worksheet->set_row( 0, undef, $format1 );           # Set format for row 1
        $worksheet->set_column( 'A:A', undef, $format2 );    # Set format for col 1
        $worksheet->write( 'A1', 'Hello' );                  # Defaults to $format1
        $worksheet->write( 'A2', 'Hello' );                  # Defaults to $format2

 
	�����������У�$hidden����Ӧ������Ϊ1���ⱻ���ڣ����磬�ڸ��Ӽ����������м䲽�裺


        $worksheet->set_column( 'D:D', 20,    $format, 1 );
        $worksheet->set_column( 'E:E', undef, undef,   1 );

    
	$level�������������еķּ���ʾ��outline level����"OUTLINES AND GROUPING IN EXCEL"�½��й��ڷּ���ʾ����������ͬ���ּ���ʾ���лᱻ���ϵ�һ����Ϊһ����һ�ķּ���ʾ��

	����������Ϊ��B��G���������˷ּ���ʾ1��

        $worksheet->set_column( 'B:G', undef, undef, 0, 1 );


	����$level����һͬʹ�õ�ʱ����$hidden����Ҳ�����������۵��ķּ���ʾ�У�

        $worksheet->set_column( 'B:G', undef, undef, 1, 1 );

  	�����۵��ķּ���ʾ����Ӧ��ʹ�ÿ�ѡ��$collapsed����ָ����һ�к����۵�����"+".

        $worksheet->set_column( 'H:H', undef, undef, 0, 0, 1 );

    �鿴��outline.pl" �� "outline_collapsed.pl" ������ȡ����ϸ��������

	Excel ��������7���ķּ���ʾ�����ˣ�$level����Ӧ���ڷ�Χ"0 <= $level <= 7"�ڡ�

  outline_settings( $visible, $symbols_below, $symbols_right, $auto_style )

	 "outline_settings()"�������ڿ���Excel�зּ���ʾ�ĳ��֡��ּ���ʾ��"OUTLINES AND GROUPING IN
    EXCEL"����������

	$visible�������ڿ��Ʒּ���ʾ�Ƿ��ɼ������ò�������Ϊ0�ᵼ�¹����������еķּ���ʾ�����ء�������ʹ��"Show Outline Symbols"���ť��������ʾ������Ĭ������Ϊ1������ʾ�ּ���

        $worksheet->outline_settings( 0 );

    The $symbols_below parameter is used to control whether the row outline
    symbol will appear above or below the outline level bar. The default
    setting is 1 for symbols to appear below the outline level bar.
	$symbols_below�������ڿ����зּ���ʾ��־���Ƿ��������ڷּ���ʾ���������Ϸ������档
	Ĭ�ϵ�����Ϊ1������ʶ�������ڷּ���ʾ�����������档

 
	"symbols_right"�������ڿ����зּ���ʾ��ʶ���Ƿ��������ڷּ���ʾ���������������Ҳࡣ
	Ĭ������Ϊ1������ʶ�������ڷּ���ʾ���ұߡ�

    The $auto_style parameter is used to control whether the automatic
    outline generator in Excel uses automatic styles when creating an
    outline. This has no effect on a file generated by "Excel::Writer::XLSX"
    but it does have an effect on how the worksheet behaves after it is
    created. The default setting is 0 for "Automatic Styles" to be turned
    off.
	$auto_style �������ڿ�����Excel�е��Զ��ּ���ʾ�������Ƿ�ʹ���Զ����񴴽��ּ���ʾ��
	��"Excel::Writer::XLSX"���ɵ��ļ���û�����𣬵��Ƕ��ڴ��������������α������������ġ�
	Ĭ������Ϊ0�����ر�"Automatic Styles" 
    The default settings for all of these parameters correspond to Excel's
    default parameters.
    �������ֲ�����Ĭ��������Excel��Ĭ�ϲ����йء�

	�� "outline_settings()"�������ƵĹ�������������ʹ�á�

  freeze_panes( $row, $col, $top_row, $left_col )  #���ᴰ��
   
	�÷������ڽ�����������Ϊˮƽ����ֱ�Ľ����������������Ҷ�����Щ������ʹ�ָ������ɼ�������Excel�е�"����->���ᴰ��"�˵�������������ͬ��

    The parameters $row and $col are used to specify the location of the
    split. It should be noted that the split is specified at the top or left
    of a cell and that the method uses zero based indexing. Therefore to
    freeze the first row of a worksheet it is necessary to specify the split
    at row 2 (which is 1 as the zero-based index). This might lead you to
    think that you are using a 1 based index but this is not the case.
	$row �� $col��������ָ���ָ���λ�á�Ӧ��ע�����Ƿָ��ɵ�Ԫ���Ķ���������ָ�������Ҹ÷���ʹ�û���0��ʼ�����������ˣ����Ṥ�����ĵ�һ�У�ָ���ڶ��У���Ϊ����0������ʱ��1�����зָ����б�Ҫ�ġ������ܵ�������Ϊ��ʹ�û���1�������������ⲻ�����⡣

    You can set one of the $row and $col parameters as zero if you do not
    want either a vertical or horizontal split.
	�����㲻��Ҫˮƽ����ֱ�ָ��������Խ�$row��$col�����е�һ������Ϊ0��

    ����:

        $worksheet->freeze_panes( 1, 0 );    # Freeze the first row
        $worksheet->freeze_panes( 'A2' );    # Same using A1 notation
        $worksheet->freeze_panes( 0, 1 );    # Freeze the first column
        $worksheet->freeze_panes( 'B1' );    # Same using A1 notation
        $worksheet->freeze_panes( 1, 2 );    # Freeze first row and first 2 columns
        $worksheet->freeze_panes( 'C2' );    # Same using A1 notation

    The parameters $top_row and $left_col are optional. They are used to
    specify the top-most or left-most visible row or column in the scrolling
    region of the panes. ���磬 to freeze the first row and to have the
    scrolling region begin at row twenty:
	$top_row �� $left_col�����ǿ�ѡ�ġ����������ڴ����Ĺ���������ָ���ɼ�����˻������˵��л��С�
	���磬������һ�в��ù��������ӵ�20�п�ʼ��

        $worksheet->freeze_panes( 1, 0, 20, 0 );

	����$top_row �� $left_col������������ʹ��A1��ʾ����

    �鿴��"panes.pl"������

  split_panes( $y, $x, $top_row, $left_col )
    This method can be used to divide a worksheet into horizontal or
    vertical regions known as panes. This method is different from the
    "freeze_panes()" method in that the splits between the panes will be
    visible to the user and each pane will have its own scroll bars.
	�÷������ڽ�����������Ϊ����������ˮƽ�Ļ���ֱ�����򡣸÷�����ͬ��"freeze_panes()"������
	���ָ��Ĵ������û��ǿɼ��ģ�����ÿ���������������Լ��Ĺ�������

    The parameters $y and $x are used to specify the vertical and horizontal
    position of the split. The units for $y and $x are the same as those
    used by Excel to specify row height and column width. However, the
    vertical and horizontal units are different from each other. Therefore
    you must specify the $y and $x parameters in terms of the row heights
    and column widths that you have set or the default values which are 15
    for a row and 8.43 for a column.
	$y �� $x �������ڷָ���ˮƽ�ʹ�ֱλ�á� $y��$x�ĵ�λ��Excelʹ�õĵ�λһ��������ָ���иߺ��п���Ȼ����ˮƽ�ĺʹ�ֱ�ĵ�λ��һ���� 
	���ˣ������밴�������úõ��иߺ��п���ָ��$y �� $x ���������ߣ�ʹ��Ĭ��ֵ���и�15,�п�8.43.

    You can set one of the $y and $x parameters as zero if you do not want
    either a vertical or horizontal split. The parameters $top_row and
    $left_col are optional. They are used to specify the top-most or
    left-most visible row or column in the bottom-right pane.
	�����㲻��ˮƽ�ָ�����ֱ�ָ��������Խ�$y �� $x����֮һ����Ϊ0��top_row ��$left_col�����ǿ�ѡ�ģ���������ָ�����ҵײ������е����ϻ������Ŀɼ��л��С�

    ����:

        $worksheet->split_panes( 15, 0,   );    # First row
        $worksheet->split_panes( 0,  8.43 );    # First column
        $worksheet->split_panes( 15, 8.43 );    # First row and column

    �÷�������ʹ��A1��ʾ����

    �鿴 "freeze_panes()"������ "panes.pl"������

  merge_range( $first_row, $first_col, $last_row, $last_col, $token, $format )
    The "merge_range()" method allows you merge cells that contain other
    types of alignment in addition to the merging:
	"merge_range()"�����������ϲ������������뷽ʽ�����˺ϲ����ĵ�Ԫ����

        my $format = $workbook->add_format(
            border => 6,
            valign => 'vcenter',
            align  => 'center',
        );

        $worksheet->merge_range( 'B3:D4', 'Vertical and horizontal', $format );

	"merge_range()"����ʹ�ù�������"write()"����д������$token���������ˣ����ᰴҪ���������֣��ַ�������ʽ��urls����������ָ��Ҫ����"write_*"������ʹ�� "merge_range_type()"�����������档

	�鿴"merge3.pl"��"merge6.pl"��ȡ�÷�����ȫ����Ϣ��

  merge_range_type( $type, $first_row, $first_col, $last_row, $last_col, ... )
    The "merge_range()" method, see above, uses "write()" to insert the
    required data into to a merged range. However, there may be times where
    this isn't what you require so as an alternative the "merge_range_type
    ()" method allows you to specify the type of data you wish to write. For
    example:
	"merge_range()"�����������棬ʹ��"write()"���ϲ��������в�����Ҫ�����ݡ�Ȼ������ʱ�����ܲ�������Ҫ�ģ�������Ϊѡ����"merge_range_type()"����������ָ������д�����������͡����磺

        $worksheet->merge_range_type( 'number',  'B2:C2', 123,    $format1 );
        $worksheet->merge_range_type( 'string',  'B4:C4', 'foo',  $format2 );
        $worksheet->merge_range_type( 'formula', 'B6:C6', '=1+2', $format3 );

    The $type must be one of the following, which corresponds to a
    "write_*()" method:
	$type������������֮һ���� "write_*()"�������ԣ�

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
	�κ����ⷶΧ֮���Ĳ���Ӧ�����κκ��ʵķ����ɽ��ܵģ�

        $worksheet->merge_range_type( 'rich_string', 'B8:C8',
                                      'This is ', $bold, 'bold', $format4 );

    Note, you must always pass a $format object as an argument, even if it
    is a default format.
    ע�⣬������һֱ����һ��$format������Ϊ��������ʹ����һ��Ĭ�ϵĸ�ʽ��
  set_zoom( $scale )
    Set the worksheet zoom factor in the range "10 <= $scale <= 400":
	�ڷ�Χ"10 <= $scale <= 400"�����ù�����������������

        $worksheet1->set_zoom( 50 );
        $worksheet2->set_zoom( 75 );
        $worksheet3->set_zoom( 300 );
        $worksheet4->set_zoom( 400 );

    The default zoom factor is 100. You cannot zoom to "Selection" because
    it is calculated by Excel at run-time.
	Ĭ�ϵ�����������100.�㲻�ܶ�ѡȡ�������ţ���Ϊ��������ʱ��Excel���㡣

    Note, "set_zoom()" does not affect the scale of the printed page. For
    that you should use "set_print_scale()".
	ע�⣬"set_zoom()" ��Ӱ����ӡҳ�ĳߴ硣���ڴˣ���Ӧ��ʹ��"set_print_scale()".

  right_to_left()
    The "right_to_left()" method is used to change the default direction of
    the worksheet from left-to-right, with the A1 cell in the top left, to
    right-to-left, with the he A1 cell in the top right.
	"right_to_left()"�������ڸı乤������Ĭ�Ϸ������ɴ������ң���A1��Ԫ�������Ϸ�����Ϊ���ҵ��󣬼�A1��Ԫ�������Ϸ���

        $worksheet->right_to_left();

    This is useful when creating Arabic, Hebrew or other near or far eastern
    worksheets that use right-to-left as the default direction.
	�������������ġ�ϣ�����Ļ������ӽ�������Զ����Ĭ��ʹ�ô��ҵ��������Ĺ�����ʱ���á�
	

  hide_zero()
   
	"hide_zero()"�������������κγ����ڵ�Ԫ���е�0ֵ��

        $worksheet->hide_zero();

	��Excel�У���ѡ�����ڹ���->ѡ��->�鿴�˵����ҵ���

  set_tab_color()
    The "set_tab_color()" method is used to change the colour of the
    worksheet tab. This feature is only available in Excel 2002 and later.
    You can use one of the standard colour names provided by the Format
    object or a colour index. See "COLOURS IN EXCEL" and the
    "set_custom_color()" method.
	"set_tab_color()"�������ڸı乤����������ɫ���ù���ֻ��Excel 2002���Ժ����á�������ʹ�ø�ʽ��������ɫ�����ṩ�ı�׼��ɫ��֮һ���鿴"COLOURS IN EXCEL" ��"set_custom_color()"������

        $worksheet1->set_tab_color( 'red' );
        $worksheet2->set_tab_color( 0x0C );

    �鿴"tab_colors.pl" ������

  autofilter( $first_row, $first_col, $last_row, $last_col )
    This method allows an autofilter to be added to a worksheet. An
    autofilter is a way of adding drop down lists to the headers of a 2D
    range of worksheet data. This is turn allow users to filter the data
    based on simple criteria so that some data is shown and some is hidden.
	�÷���������������������һ���Զ�ɸѡ���ܡ�

    To add an autofilter to a worksheet:
	������������һ���Զ�ɸѡ��

        $worksheet->autofilter( 0, 0, 10, 3 );
        $worksheet->autofilter( 'A1:D11' );    # Same as above in A1 notation.

    Filter conditions can be applied using the "filter_column()" or
    "filter_column_list()" method.
	ɸѡ������ʹ��"filter_column()"��"filter_column_list()"����Ӧ�á�

    �鿴"autofilter.pl" ������

  filter_column( $column, $expression )
    The "filter_column" method can be used to filter columns in a autofilter
    range based on simple conditions.
	"filter_column"���������ڸ��ݼ򵥵�������һ���Զ�ɸѡ��Χ�ڹ����С�

    NOTE: It isn't sufficient to just specify the filter condition. You must
    also hide any rows that don't match the filter condition. Rows are
    hidden using the "set_row()" "visible" parameter. "Excel::Writer::XLSX"
    cannot do this automatically since it isn't part of the file format. See
    the "autofilter.pl" program in the examples directory of the distro for
    an example.
	ע�⣺����ָ�����������ǲ����ġ���Ҳ���������κβ�ƥ�������������С�
	ʹ��"set_row()" "visible" �����������С�
	"Excel::Writer::XLSX"�����Զ�������������Ϊ�������ļ���ʽ��һ���֡��鿴"autofilter.pl" ������

    The conditions for the filter are specified using simple expressions:
	ʹ�ü򵥵ı���ʽָ������������

        $worksheet->filter_column( 'A', 'x > 2000' );
        $worksheet->filter_column( 'B', 'x > 2000 and x < 5000' );

    The $column parameter can either be a zero indexed column number or a
    string column name.
	$column���������Ǵ�0������һ���б��Ż�һ���ַ���������

	�����Ĳ������ǿ��õģ�

        ������        ͬ����
           ==           =   eq  =~
           !=           <>  ne  !=
           >
           <
           >=
           <=

           and          &&
           or           ||

	��������ͬ���ʽ�����������������ʹ�ñ���ʽ���﷨�ǡ���һ������Ҫ������ʽ�ᱻExcel���Ͷ�����perl

	һ������ʽ����һ����һ���������ɻ���"and" �� "or"�������ֿ���2���������ɣ����磺

        'x <  2000'
        'x >  2000'
        'x == 2000'
        'x >  2000 and x <  5000'
        'x == 2000 or  x == 5000'

 
	�ڱ���ʽ��ʹ��"Blanks"�� "NonBlanks"ֵ�ܴﵽ���˿տ����ǿհ����ݵ����ã�

        'x == Blanks'
        'x == NonBlanks'

	ExcelҲ����һЩ�򵥵��ַ���ƥ��������

        'x =~ b*'   # begins with b
        'x !~ b*'   # doesn't begin with b
        'x =~ *b'   # ends with b
        'x !~ *b'   # doesn't end with b
        'x =~ *b*'  # contains b
        'x !~ *b*'  # doesn't contains b


	������ʹ��"*"ƥ�������ַ������֣�ʹ��"?"ƥ����һ�ַ������֡�Excel�Ĺ�������֧����������������ʽ���ʡ�Excel����������ʽ�ַ���ʹ��"~"����ת�塣

	������ռλ������"x"�ܱ������ļ����ַ������档ʵ�ʵ�ռλ�������ڲ������ԣ������������ǵȼ۵ģ�

        'x     < 2000'
        'col   < 2000'
        'Price < 2000'

    Also, note that a filter condition can only be applied to a column in a
    range specified by the "autofilter()" Worksheet method.
	ע�⣬������������Ӧ�õ���"autofilter()"������������ָ����Χ�����ϡ�

    �鿴"autofilter.pl" ������


	ע��Spreadsheet::WriteExcel֧������10�����͵Ĺ��ˡ���ЩĿǰ���� Excel::Writer::XLSX֧�֣��������Ժ����ӡ�

  filter_column_list( $column, @matches )

	��Excel 2007��ǰֻ��1��2��������������������չʾ��"filter_column" ������

    Excel 2007 introduced a new list style filter where it is possible to
    specify 1 or more 'or' style criteria. ���磬 if your column
    contained data for the first six months  Then if you selected
    'March', 'April' and 'May' they would be displayed as shown on the
    right.
	Excel 2007����һ���µ��б����͹��ˣ���ָ��1��������'or'���͵ı�׼�����磬�����������к���ǰ�����µ����ݣ���ʼ���ݻᰴ����ѡ��the initial data would be displayed as all selected as shown on the left.��������ѡ���� 'March', 'April' �� 'May' ���ǻ���ʾ���ұߡ�


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
	"filter_column_list()" ���������ڴ�����Щ���͵Ĺ��ˣ�

        $worksheet->filter_column_list( 'A', 'March', 'April', 'May' );

    The $column parameter can either be a zero indexed column number or a
    string column name.
	$column ���������Ǵ�0�������б��Ż�һ���ַ���������

    ����ѡ��һ����������׼:

        $worksheet->filter_column_list( 0, 'March' );
        $worksheet->filter_column_list( 1, 100, 110, 120, 130 );

    NOTE: It isn't sufficient to just specify the filter condition. You must
    also hide any rows that don't match the filter condition. Rows are
    hidden using the "set_row()" "visible" parameter. "Excel::Writer::XLSX"
    cannot do this automatically since it isn't part of the file format. See
    the "autofilter.pl" program in the examples directory of the distro for
    an example. e conditions for the filter are specified using simple
    expressions:
	ע�⣺����ָ�����������ǲ����ġ���Ҳ���������κβ�ƥ�������������С�
	ʹ��"set_row()" "visible" �����������С�
	"Excel::Writer::XLSX"�����Զ�������������Ϊ�������ļ���ʽ��һ���֡��鿴"autofilter.pl" ������

  convert_date_time( $date_string )
  
	"convert_date_time()" �������ڲ���"write_date_time()" ����ʹ�ã����ڽ������ַ���ת��Ϊ��Excel�д������ں�ʱ�������֡�

    Ϊ��ʵ��Ŀ�ģ�����Ϊһ�ֹ���������¶��������ǰ��
	$date_string��ʽ��"write_date_time()"����������ϸ˵����

PAGE SET-UP METHODS
   ��ӡ��ʱ�ҳ��set-up����Ӱ��һ�Ź����������Ρ����ǿ���������ҳü��ҳ�ź�ҳ�߾๦�ܡ���Щ�������Ǳ�׼�Ĺ�����������Ϊ���������������õ������½���˵�����ǵ�ʹ�á�
	�����ķ�������ҳ�������ǿ��õģ�

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

    	��ʹ��Excel::Writer::XLSX����ʱ��ͨ���������ǽ�ͬһ��ҳ����������Ӧ�õ��������е����й������С�������ʹ��"workbook"����"sheets()"����ͨ�����ʹ������еĹ��������������ɣ�

        for $worksheet ( $workbook->sheets() ) {
            $worksheet->set_landscape();
        }

  set_landscape()
    This method is used to set the orientation of a worksheet's printed page
    to landscape:
	

        $worksheet->set_landscape();    # Landscape mode

  set_portrait() #�������Ÿ�ʽ����ӡˢ��(��ҳ����ͼ��������)���Ÿ�ʽ��
    This method is used to set the orientation of a worksheet's printed page
    to portrait. The default worksheet orientation is portrait, so you won't
    generally need to call this method.
	�÷����������ù�������ӡҳ���������ŵķ�����Ĭ�ϵĹ��������������ţ�����ͨ���㲻��Ҫ���ø÷�����

        $worksheet->set_portrait();    # Portrait mode

  set_page_view()
	�÷���������"ҳ���鿴/����"ģʽ��ʾ��������

        $worksheet->set_page_view();

  set_paper( $index )
    This method is used to set the paper format for the printed output of a
    worksheet. The following paper styles are available:
	�÷�������Ϊ�������Ĵ�ӡ��������ҳ����ʽ�������ǿ��õ�ֽ�����ͣ�

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
	ע�⣬�������е�ֽ�����Ͷ����ն��û����ǿ��õģ���Ϊ���������û��Ĵ�ӡ��֧�ֵ�ҳ����ʽ�����ˣ�����ʹ�ñ�׼��ֽ�����͡�

        $worksheet->set_paper( 1 );    # US Letter
        $worksheet->set_paper( 9 );    # A4

 
	������û��ָ��ֽ�����ͣ���������ʹ�ô�ӡ��Ĭ�ϵ�ֽ�Ŵ�ӡ��

  center_horizontally()
    Center the worksheet data horizontally between the margins on the
    printed page:
	�ڴ�ӡҳ����ҳ�߾�֮��ˮƽ���ж��빤�������ݣ�

        $worksheet->center_horizontally();

  center_vertically()
    Center the worksheet data vertically between the margins on the printed
    page:
	�ڴ�ӡҳ��ҳ�߾�֮�䴹ֱ���ж��빤�������ݣ�

        $worksheet->center_vertically();

  set_margins( $inches )
    There are several methods available for setting the worksheet margins on
    the printed page:
	�м��ֿ��õķ����������ô�ӡҳ���Ĺ�����ҳ�߾ࣺ

        set_margins()        # ������ҳ�߾���Ϊͬ����ֵ
        set_margins_LR()     # ����ҳ�߾�����ҳ�߾���Ϊͬ����ֵ
        set_margins_TB()     # ����ҳ�߾�����ҳ�߾���Ϊͬ����ֵ
        set_margin_left();   # ������ҳ�߾�
        set_margin_right();  # ������ҳ�߾�
        set_margin_top();    # Set top margin������ҳ�߾�
        set_margin_bottom(); # Set bottom margin������ҳ�߾�

    All of these methods take a distance in inches as a parameter. Note: 1
    inch = 25.4mm. ";-)" The default left and right margin is 0.7 inch. The
    default top and bottom margin is 0.75 inch. Note, these defaults are
    different from the defaults used in the binary file format by
    Spreadsheet::WriteExcel.
	������Щ������Ӣ��������Ϊ������ע�⣺1Ӣ��=25.4���ס�Ĭ�ϵ���ҳ�߾�����ҳ�߾���0.7Ӣ�硣
	Ĭ�ϵ���ҳ�߾�����ҳ�߾���0.75Ӣ�硣ע�⣬��ЩĬ��ֵ��Spreadsheet::WriteExcel��ʹ�õĶ������ļ���ʽ��Ĭ��ֵ��ͬ��

  set_header( $string, $margin )
   
	ҳü��ҳ��ʹ��$string���ɣ�$string����ͨ�ı��Ϳ����ַ����ɡ� $margin�����ǿ�ѡ�ġ�

	���õĿ����ַ��ǣ�

        �����ַ�            ����                ����
        =======             ========            ===========
        &L                  ����                ������
        &C                                      ���ж���
        &R                                      �Ҷ���

        &P                  ��Ϣ                ҳ��
        &N                                      ��ҳ��
        &D                                      ����
        &T                                      ʱ��
        &F                                      �ļ���
        &A                                      ��������
        &Z                                      ������·��

        &fontsize           ����                ������С
        &"font,style"                           ����������������
        &U                                      ���»���
        &E                                      ˫�»���
        &S                                      ɾ����
        &X                                      �ϱ�
        &Y                                      �±�

        &&                  ����                ��������&

	ͨ�����ı�ǰ��ǰ�ÿ����ַ�&L��&C����&R,���Խ�ҳü��ҳ���е��ı����������룩Ϊ���󡢾��к��Ҷ��롣

    ���磬 (ʹ�� ASCII ��ͼ��ʾ����):

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
	���ڴ��ı���������û��ָ���κζ��뷽ʽ���ı������ж��롣Ȼ����������ָ���������������κθ�ʽ�����������ı�ǰǰ��&C���š�

        $worksheet->set_header('Hello');

         ---------------------------------------------------------------
        |                                                               |
        |                          Hello                                |
        |                                                               |

	��������ÿ�����������ж����ı���

        $worksheet->set_header('&LCiao&CBello&RCielo');

         ---------------------------------------------------------------
        |                                                               |
        | Ciao                     Bello                          Cielo |
        |                                                               |

	���������������������仯ʱ����Ϣ�����ַ���ΪExcel�����µı�����ʱ��������ʹ���û�Ĭ�ϵĸ�ʽ��

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


	������ͨ��������ǰǰ�ÿ����ַ�&n,"n"��������С����ָ���ı�������������С��

        $worksheet1->set_header( '&C&30Hello Big' );
        $worksheet2->set_header( '&C&10Hello Small' );


	���������ı�ǰǰ�ÿ�������&"font,style"��ָ���ı����������塣"font"������"Courier New" �� "Times New Roman"����������"style"�Ǳ�׼��Windows��������֮һ��"Regular", "Italic", "Bold" �� "Bold Italic":

        $worksheet1->set_header( '&C&"Courier New,Italic"Hello' );
        $worksheet2->set_header( '&C&"Courier New,Bold Italic"Hello' );
        $worksheet3->set_header( '&C&"Times New Roman,Regular"Hello' );

	��������Щ�������������������ӵ�ҳü��ҳ���ǿ��ܵġ���Ϊ���ڽ�������ҳü��ҳ�ŵİ�������������Excel�м�¼һ��ҳ�����õĺ꣬���Ҳ鿴VBA�����ĸ�ʽ�ַ�������סVBAʹ��2��˫����""��������˫����".��������������һ�����ӣ��ȼ۵�VBA���뿴������������

        .LeftHeader   = ""
        .CenterHeader = "&""Times New Roman,Regular""Hello"
        .RightHeader  = ""

	��Ӧ��ʹ��2��and����"&&"��ҳü��ҳ���б�ʾһ������and����"&":

        $worksheet1->set_header('&CCuriouser && Curiouser - Attorneys at Law');

    As stated above the margin parameter is optional. As with the other
    margins the value should be in inches. The default header and footer
    margin is 0.3 inch. Note, the default margin is different from the
    default used in the binary file format by Spreadsheet::WriteExcel. The
    header and footer margin size can be set as follows:
     �����������У���ʼҳ�߾������ǿ�ѡ�ġ���������ҳ�߾���ֵӦ����Ӣ�硣Ĭ�ϵ�ҳü��ҳ��ҳ�߾���0.3Ӣ�硣ע�⣬��ЩĬ��ҳ�߾�ֵ��Spreadsheet::WriteExcel��ʹ�õĶ������ļ���ʽ��Ĭ��ֵ��ͬ��ҳü��ҳ�ŵ�ҳ�߾���С�����������£�
        $worksheet->set_header( '&CHello', 0.75 );

    The header and footer margins are independent of the top and bottom
    margins.
	ҳü��ҳ�ŵ�ҳ�߾������ڶ����͵ײ���ҳ�߾ࡣ

  
	ע�⣬ҳü��ҳ���ַ�����������255���ַ�������255�����ַ��������ᱻд�벢����һ�����档

 
	"set_header()"����Ҳ�ܴ���UTF-8��ʽ��Unicode�ַ�����

        $worksheet->set_header( "&C\x{263a}" )

    �鿴 "headers.pl" ������

  set_footer()
  
    "set_footer()" ������"set_header()"�������﷨һ���������档
  repeat_rows( $first_row, $last_row )
    Set the number of rows to repeat at the top of each printed page.
	��ÿ�Ŵ�ӡҳ�Ķ���������Ҫ���Ƶ�������

    For large Excel documents it is often desirable to have the first row or
    rows of the worksheet print out at the top of each page. This can be
    achieved by using the "repeat_rows()" method. The parameters $first_row
    and $last_row are zero based. The $last_row parameter is optional if you
    only wish to specify one row:
	���ںܴ���Excel�ļ�����ÿҳ�Ķ�����ӡ�������ĵ�һ�л�ǰ����ͨ����ֵ�õġ�������ʹ�� "repeat_rows()" ����������$first_row��$last_row�����ǻ���0�ġ� ������ֻ��ָ��һ�У�$last_row�����ǿ�ѡ�ģ�

        $worksheet1->repeat_rows( 0 );    # ���Ƶ�һ��
        $worksheet2->repeat_rows( 0, 1 ); # ����ǰ2��

  repeat_columns( $first_col, $last_col )
    Set the columns to repeat at the left hand side of each printed page.
	��ÿ�Ŵ�ӡҳ������������Ҫ���Ƶ�������

    For large Excel documents it is often desirable to have the first column
    or columns of the worksheet print out at the left hand side of each
    page. This can be achieved by using the "repeat_columns()" method. The
    parameters $first_column and $last_column are zero based. The
    $last_column parameter is optional if you only wish to specify one
    column. You can also specify the columns using A1 column notation, see
    the note about "Cell notation".
	
    ���ںܴ���Excel�ļ�����ÿҳ��������ӡ�������ĵ�һ�л�ǰ����ͨ����ֵ�õġ�������ʹ�� "repeat_columns()" ����������$first_column��$last_column�����ǻ���0�ġ� ������ֻ��ָ��һ�У�$last_column�����ǿ�ѡ�ġ�������ʹ��A1�б�ʾ��ָ��������

        $worksheet1->repeat_columns( 0 );        # Repeat the first column
        $worksheet2->repeat_columns( 0, 1 );     # Repeat the first two columns
        $worksheet3->repeat_columns( 'A:A' );    # Repeat the first column
        $worksheet4->repeat_columns( 'A:B' );    # Repeat the first two columns

  hide_gridlines( $option )
    This method is used to hide the gridlines on the screen and printed
    page. Gridlines are the lines that divide the cells on a worksheet.
     If you have defined your own cell borders you may wish to
    hide the default gridlines.
	�÷�������������Ļ�ϵ������ߺʹ�ӡ����ҳ�档���������ڹ������зָ���Ԫ�����ߡ�
	Screen and printed gridlines are turned on by default in an Excel
    worksheet.
   �����㶨�������Լ��ĵ�Ԫ���߿�������������Ĭ�ϵ������ߡ�
        $worksheet->hide_gridlines();

	������$optionֵ����Ч�ģ�

        0 : ������������
        1 : ֻ���ش�ӡ����������Hide printed gridlines only
        2 : ������Ļ�ʹ�ӡ����������Hide screen and printed gridlines

    If you don't supply an argument or use "undef" the default option is 1,
    i.e. only the printed gridlines are hidden.
	������û���ṩ������ʹ��"undef"����Ĭ�ϵ�ѡ����1��i.e��ֻ�д�ӡ���������߱����ء�

  print_row_col_headers()
    Set the option to print the row and column headers on the printed page.
	�ڴ�ӡҳ��������ѡ���Դ�ӡ�б������б��⡣

    An Excel worksheet looks something like the following��
	һ�Ź�������������������������

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
	�������ǹ�������������������ĸ�����֡���Ϊ��Щ������Ҫ�ڹ�������ָ��λ�ã�����һ�㲻�������ڴ�ӡҳ���ϡ���������������һ����ӡ��������ʹ��"print_row_col_headers()"������

        $worksheet->print_row_col_headers();

    Do not confuse these headers with page headers as described in the
    "set_header()" section above.
	��Ҫ����Щ�����������ᵽ���й�ҳ��������"set_header()"�½�Ū������

  print_area( $first_row, $first_col, $last_row, $last_col )
    This method is used to specify the area of the worksheet that will be
    printed. All four parameters must be specified. You can also use A1
    notation, 
	�÷�������ָ��������ӡ�Ĺ�����������4���������붼ָ������Ҳ����ʹ��A1��ʾ����

        $worksheet1->print_area( 'A1:H20' );    # Cells A1 to H20
        $worksheet2->print_area( 0, 0, 19, 7 ); # The same
        $worksheet2->print_area( 'A:H' );       # Columns A to H if rows have data

  print_across()
    The "print_across" method is used to change the default print direction.
    This is referred to by Excel as the sheet "page order".
	"print_across"�������ڸı�Ĭ�ϵĴ�ӡ����������Excel�������б���Ϊ��ҳ��˳�򡱡�

        $worksheet->print_across();

    The default page order is shown below for a worksheet that extends over
    4 pages. The order is called "down then across":
	������ʾ����ӵ�г���4ҳ�Ĺ�������Ĭ�ϵ�ҳ��˳�򡣸�˳������������Ȼ�󽻲桱

        [1] [3]
        [2] [4]

    However, by using the "print_across" method the print order will be
    changed to "across then down":
	Ȼ����ͨ��ʹ��"print_across"��������ӡ˳��������Ϊ"��������"��

        [1] [2]
        [3] [4]

  fit_to_pages( $width, $height )
    The "fit_to_pages()" method is used to fit the printed area to a
    specific number of pages both vertically and horizontally. If the
    printed area exceeds the specified number of pages it will be scaled
    down to fit. This guarantees that the printed area will always appear on
    the specified number of pages even if the page size or margins change.
	"fit_to_pages()"�������ڴ�ֱ��ˮƽ��ʹ��ӡ������ָ��ҳ�����ʡ�������ӡ���򳬹���ָ����ҳ�������ᰴ������С����Ӧ���Ᵽ֤�˼�ʹҳ���ߴ���ҳ�߾෢���仯����ӡ����Ҳ��һֱ������ָ��ҳ�ϡ�

        $worksheet1->fit_to_pages( 1, 1 );    # Fit to 1x1 pages
        $worksheet2->fit_to_pages( 2, 1 );    # Fit to 2x1 pages
        $worksheet3->fit_to_pages( 1, 2 );    # Fit to 1x2 pages

	��ӡ������ʹ������������"print_area()"�������塣

    A common requirement is to fit the printed output to *n* pages wide but
    have the height be as long as necessary. To achieve this set the $height
    to zero:
    ͨ���������ǽ���ӡ����Ϊnҳ�������ø߶Ⱦ����ܵس������԰� $height����Ϊ0���ﵽҪ����
        $worksheet1->fit_to_pages( 1, 0 );    # 1 page wide and as long as necessary

    Note that although it is valid to use both "fit_to_pages()" and
    "set_print_scale()" on the same worksheet only one of these options can
    be active at a time. The last method call made will set the active
    option.
	ע�⣬������ͬһ�Ź�������ʹ��"fit_to_pages()" �� "set_print_scale()" ����ȷ�ģ���һ��ֻ�ܼ������е�һ��ѡ�����һ���������û����ü���ѡ�

    Note that "fit_to_pages()" will override any manual page breaks that are
    defined in the worksheet.
	 ע��"fit_to_pages()"����д�κ��ֲ�ҳ

  set_start_page( $start_page )
    The "set_start_page()" method is used to set the number of the starting
    page when the worksheet is printed out. The default value is 1.
	 "set_start_page()"�����������ù�������ӡʱ����ʼҳ��Ĭ��ֵ��1.

        $worksheet->set_start_page( 2 );

  set_print_scale( $scale )
    Set the scale factor of the printed page. Scale factors in the range "10
    <= $scale <= 400" are valid:
	���ô�ӡҳ�ı���ϵ�����ڷ�Χ"10 <= $scale <= 400"�ڵı���ϵ������Ч�ģ�

        $worksheet1->set_print_scale( 50 );
        $worksheet2->set_print_scale( 75 );
        $worksheet3->set_print_scale( 300 );
        $worksheet4->set_print_scale( 400 );

    The default scale factor is 100. Note, "set_print_scale()" does not
    affect the scale of the visible page in Excel. For that you should use
    "set_zoom()".
	Ĭ�ϵı���ϵ����100.ע�⣬"set_print_scale()"��Ӱ��Excel�ɼ�ҳ�ĳߴ硣���ڴˡ���Ӧʹ��"set_zoom()"��

    Note also that although it is valid to use both "fit_to_pages()" and
    "set_print_scale()" on the same worksheet only one of these options can
    be active at a time. The last method call made will set the active
    option.
	ҲҪע�⣬������ͬһ�Ź�������ʹ��"fit_to_pages()" �� "set_print_scale()" ����ȷ�ģ���һ��ֻ�ܼ������е�һ��ѡ�����һ���������û����ü���ѡ�

  set_h_pagebreaks( @breaks )
    Add horizontal page breaks to a worksheet. A page break causes all the
    data that follows it to be printed on the next page. Horizontal page
    breaks act between rows. To create a page break between rows 20 and 21
    you must specify the break at row 21. However in zero index notation
    this is actually row 20. So you can pretend for a small while that you
    are using 1 index notation:
	����ˮƽ��ҳ�����������С���ҳ����������������������������һҳ�б���ӡ��ˮƽ��ҳ������֮�������á�Ϊ�ڵ�20�к͵�21��֮�䴴����ҳ�����������ڵ�21��ָ����ҳ��Ȼ��������0��ʼ�����ı�ʾ���У���ʵ�����ǵ�20�С����������Լ�װ����ʹ��1������ʾ����

        $worksheet1->set_h_pagebreaks( 20 );    # Break between row 20 and 21

    The "set_h_pagebreaks()" method will accept a list of page breaks and
    you can call it more than once:
	"set_h_pagebreaks()"����������һ�зָ������������Զ��ε��ø÷�����

        $worksheet2->set_h_pagebreaks( 20,  40,  60,  80,  100 );    # Add breaks
        $worksheet2->set_h_pagebreaks( 120, 140, 160, 180, 200 );    # Add some more

    Note: If you specify the "fit to page" option via the "fit_to_pages()"
    method it will override all manual page breaks.
	ע�⣬������ͨ�� "fit_to_pages()"����ָ����"fit to page"ѡ����Ḳ�����е��ֶ���ҳ����

    There is a silent limitation of about 1000 horizontal page breaks per
    worksheet in line with an Excel internal limitation.
	��Excel�ڲ�����һ����ÿ�Ź�������ˮƽ��ҳ������Ϊ1000����

  set_v_pagebreaks( @breaks )
    Add vertical page breaks to a worksheet. A page break causes all the
    data that follows it to be printed on the next page. Vertical page
    breaks act between columns. To create a page break between columns 20
    and 21 you must specify the break at column 21. However in zero index
    notation this is actually column 20. So you can pretend for a small
    while that you are using 1 index notation:
	���Ӵ�ֱ��ҳ�����������С���ҳ����������������������������һҳ�б���ӡ����ֱ��ҳ������֮�������á�Ϊ�ڵ�20�к͵�21��֮�䴴����ҳ�����������ڵ�21��ָ����ҳ��Ȼ��������0��ʼ�����ı�ʾ���У���ʵ�����ǵ�20�С����������Լ�װ����ʹ��1������ʾ����

        $worksheet1->set_v_pagebreaks(20); #��20��21��֮����ҳ

    The "set_v_pagebreaks()" method will accept a list of page breaks and
    you can call it more than once:
	"set_v_pagebreaks()"����������һ�зָ������������Զ��ε��ø÷�����


        $worksheet2->set_v_pagebreaks( 20,  40,  60,  80,  100 );    # Add breaks
        $worksheet2->set_v_pagebreaks( 120, 140, 160, 180, 200 );    # Add some more

    Note: If you specify the "fit to page" option via the "fit_to_pages()"
    method it will override all manual page breaks.
	ע�⣬������ͨ�� "fit_to_pages()"����ָ����"fit to page"ѡ����Ḳ�����е��ֶ���ҳ����

CELL FORMATTING   #��Ԫ����ʽ��
    This section describes the methods and properties that are available for
    formatting cells in Excel. The properties of a cell that can be
    formatted include: fonts, colours, patterns, borders, alignment and
    number formatting.
	���½�������Excel�и�ʽ����Ԫ������Щ���������Կ��á������ڸ�ʽ����Ԫ�������԰��������塢��ɫ����ʽ���߿򡢶��������ָ�ʽ����

  ������ʹ�ø�ʽ����
    Cell formatting is defined through a Format object. Format objects are
    created by calling the workbook "add_format()" method as follows:
	��Ԫ���ĸ�ʽ��ͨ����ʽ���������ġ�ͨ���������µĹ�����"add_format()"����������ʽ������

        my $format1 = $workbook->add_format();            # Set properties later
        my $format2 = $workbook->add_format( %props );    # Set at creation

    The format object holds all the formatting properties that can be
    applied to a cell, a row or a column. The process of setting these
    properties is discussed in the next section.
	��ʽ��������������Ӧ�õ���Ԫ���ĸ�ʽ���ԣ�һ�л�һ�С�����һ�½�����������Щ���ԵĲ��衣

    Once a Format object has been constructed and its properties have been
    set it can be passed as an argument to the worksheet "write" methods as
    follows:
	һ�������˸�ʽ�����������������ǵ����ԣ������԰����·�����Ϊ�������ݸ���������"write"������

        $worksheet->write( 0, 0, 'One', $format );
        $worksheet->write_string( 1, 0, 'Two', $format );
        $worksheet->write_number( 2, 0, 3, $format );
        $worksheet->write_blank( 3, 0, $format );

    Formats can also be passed to the worksheet "set_row()" and
    "set_column()" methods to define the default property for a row or
    column.
	��ʽҲ���Դ��ݸ���������"set_row()"��"set_column()"����Ϊ�л��ж���Ĭ�����ԡ�

        $worksheet->set_row( 0, 15, $format );
        $worksheet->set_column( 0, 0, 15, $format );

  Format methods and Format properties
    ��ʽ�����͸�ʽ����
    The following table shows the Excel format categories, the formatting
    properties that can be applied and the equivalent object method:
	�����ı�����ʾ��Excel�ĸ�ʽ���𣬼���ʹ�õĸ�ʽ���Ժ͵ȼ۵Ķ��󷽷���

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
    method interface or by setting the property directly. ���磬, a
    typical use of the method interface would be as follows:
	��2�з������ø�ʽ���ԣ�ʹ�ö��󷽷��ӿڻ�ֱ���������ԡ����磬�����ǵ��͵ķ����ӿ��÷���

        my $format = $workbook->add_format();
        $format->set_bold();
        $format->set_color( 'red' );

    By comparison the properties can be set directly by passing a hash of
    properties to the Format constructor:
	ͨ���Ƚϣ�����ʽ���캯������һ������ɢ����ֱ���������ԣ�

        my $format = $workbook->add_format( bold => 1, color => 'red' );

    or after the Format has been constructed by means of the
    "set_format_properties()" method as follows:
	���ڸ�ʽ����֮���������ķ���ͨ��"set_format_properties()"�����������ԣ�

        my $format = $workbook->add_format();
        $format->set_format_properties( bold => 1, color => 'red' );

    You can also store the properties in one or more named hashes and pass
    them to the required method:
	��Ҳ���Խ����Դ洢��һ������������ɢ���в�����Ҫ�ķ������ݸ����ǣ�

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
	������ϲ��ͨ�����������������ԣ��������ƿ��ܸ��á����򣬴������Ե����캯����������Щ���������ĵ���ʵ�á�ʹ������ɢ�е�һ�������ô�������������������ʾ�������еĹ���������֮�乲����ʽ��
	

    The Perl/Tk style of adding properties is also supported:
	Ҳ֧������Perl/Tk���������ԣ�

        my %font = (
            -font  => 'Arial',
            -size  => 12,
            -color => 'blue',
            -bold  => 1,
        );

  Working with formatsʹ�ø�ʽ
    The default format is Arial 10 with all other properties off.
	Ĭ�ϵĸ�ʽ��Arial 10���������Զ��رա�

    Each unique format in Excel::Writer::XLSX must have a corresponding
    Format object. It isn't possible to use a Format with a write() method
    and then redefine the Format for use at a later stage. This is because a
    Format is applied to a cell not in its current state but in its final
    state. Consider the following example:
	��Excel::Writer::XLSX �У�ÿ�������ĸ�ʽ����һ����Ӧ�ĸ�ʽ������ʹ�ô���write() �����ĸ�ʽȻ�����Ժ�ʹ�ý׶������¶�����ʽ�ǲ����еġ�
	������Ϊ��Ӧ�õ���Ԫ���еĸ�ʽ���������ǵĵ�ǰ״̬���������ǵ�����״̬���������������ӣ�

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
	��Ԫ��A1��ָ����ʽ$format,����ʼ������Ϊ��ɫ��Ȼ������ɫ����������Ϊ��ɫ����Excel��ʾ��Ԫ��A1ʱ��������ʾ��ʽ������״̬���˴�����ɫ��

    In general a method call without an argument will turn a property on,
    ���磬:
	ͨ�������������ķ������ûῪ��һ�����ԣ����磺

        my $format1 = $workbook->add_format();
        $format1->set_bold();       # Turns bold on
        $format1->set_bold( 1 );    # Also turns bold on
        $format1->set_bold( 0 );    # Turns bold off

FORMAT METHODS ��ʽ����
    The Format object methods are described in more detail in the following
    sections. In addition, there is a Perl program called "formats.pl" in
    the "examples" directory of the WriteExcel distribution. This program
    creates an Excel workbook called "formats.xlsx" which contains examples
    of almost all the format types.
	�������½���ϸ�����˸�ʽ���󷽷������⣬��һ������"formats.pl"��Perl���򡣸ó��򴴽���һ����Ϊ"formats.xlsx"��Excel���������������˼������и�ʽ���͵����ӡ�

	�����ĸ�ʽ�����ǿ��õģ�

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
	�����ķ�����ֱ���������ԡ����磬"$format->set_bold()"������ "$workbook->add_format(bold => 1)"�����ȼۡ�

  set_format_properties( %properties )
    The properties of an existing Format object can be also be set by means
    of "set_format_properties()":
	ͨ������"set_format_properties()"Ҳ������һ���Ѿ����ڵĸ�ʽ���������ԣ�

        my $format = $workbook->add_format();
        $format->set_format_properties( bold => 1, color => 'red' );

    However, this method is here mainly for legacy reasons. It is preferable
    to set the properties in the format constructor:
	Ȼ����������ʷ�����÷�����Ҫ��������ڸ�ʽ���캯�����������Ը����ʣ�

        my $format = $workbook->add_format( bold => 1, color => 'red' );

  set_font( $fontname )
        Default state:      Font is Arial
        Default action:     None
        Valid args:         Any valid font name

     ָ��ʹ�õ�����:

        $format->set_font('Times New Roman');

    Excel can only display fonts that are installed on the system that it is
    running on. Therefore it is best to use the fonts that come as standard
    such as 'Arial', 'Times New Roman' and 'Courier New'. See also the Fonts
    worksheet created by formats.pl
	Excelֻ����ʾ��װ��ϵͳ����ʹ���ŵ����塣���ˣ�����ʹ����Ϊ��׼������'Arial', 'Times New Roman' �� 'Courier New'.���塣�鿴��formats.pl�����Ĺ��������塣

  set_size()
        Default state:      Font size is 10
        Default action:     Set font size to 1
        Valid args:         Integer values from 1 to as big as your screen.

    Set the font size. Excel adjusts the height of a row to accommodate the
    largest font size in the row. You can also explicitly specify the height
    of a row using the set_row() worksheet method.
	����������С��Excel�������и�����Ӧ���е��������塣��Ҳ������ʽ��ʹ��set_row()����������ָ���иߡ�

        my $format = $workbook->add_format();
        $format->set_size( 30 );

  set_color()
        Default state:      Excels��Ĭ����ɫ��ͨ���Ǻ�ɫ
        Default action:     ����Ĭ����ɫ
        Valid args:         8..63֮�����������������ַ���:
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

    
	����������ɫ��"set_color()"�����÷����£�

        my $format = $workbook->add_format();
        $format->set_color( 'red' );
        $worksheet->write( 0, 0, 'wheelbarrow', $format );

   	ע�⣺"set_color()" �������ڵ�Ԫ������������ɫ��Ҫ���õ�Ԫ������ɫ��ʹ��"set_bg_color()" ��
    "set_pattern()" ����.

	���������ӣ����鿴formats.pl������'Named colors' �� 'Standard colors'��������

    

  set_bold()
        Default state:      bold is off
        Default action:     Turn bold on
        Valid args:         0, 1

    
	����������bold�������ԣ�

        $format->set_bold();  # Turn bold on

  set_italic()
        Default state:      Italic is off
        Default action:     Turn italic on
        Valid args:         0, 1


	����������б�����ԣ�

        $format->set_italic();  # Turn italic on

  set_underline()
        Default state:      Underline is off
        Default action:     Turn on single underline
        Valid args:         0  = û���»���
                            1  = ��һ�»���
                            2  = ˫�»���
                            33 = Single accounting underline
                            34 = Double accounting underline

	�����������»������ԡ�

        $format->set_underline();   # Single underline

  set_font_strikeout()
        Default state:      Strikeout is off
        Default action:     Turn strikeout on
        Valid args:         0, 1

	����������ɾ�������ԡ�

  set_font_script()
        Default state:      Super/Subscript is off
        Default action:     Turn Superscript on
        Valid args:         0  = Normal
                            1  = Superscript
                            2  = Subscript

   	�����������ϱ�/�±����ԡ�

  set_font_outline()
        Default state:      Outline is off
        Default action:     Turn outline on
        Valid args:         0, 1

    ��֧��Mac.

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
	�÷������ڶ���Excel�����ֵ����ָ�ʽ����������һ�������Ƿ���ʾΪ�����������������ڡ�����ֵ�������û������ĸ�ʽ��

    The numerical format of a cell can be specified by using a format string
    or an index to one of Excel's built-in formats:
	��Ԫ�������ָ�ʽ��ʹ��һ����ʽ���ַ�����Excel���ڽ���ʽ����ָ����

        my $format1 = $workbook->add_format();
        my $format2 = $workbook->add_format();
        $format1->set_num_format( 'd mmm yyyy' );    # Format string
        $format2->set_num_format( 0x0f );            # Format index

        $worksheet->write( 0, 0, 36892.521, $format1 );    # 1 Jan 2001
        $worksheet->write( 0, 0, 36892.521, $format2 );    # 1-Jan-01

   	ʹ�ø�ʽ���ַ����ܶ����ܸ��ӵ����ָ�ʽ.

        $format01->set_num_format( '0.000' );
        $worksheet->write( 0, 0, 3.1415926, $format01 );    # 3.142

        $format02->set_num_format( '#,##0' );
        $worksheet->write( 1, 0, 1234.56, $format02 );      # 1,235

        $format03->set_num_format( '#,##0.00' );
        $worksheet->write( 2, 0, 1234.56, $format03 );      # 1,234.56

        $format04->set_num_format( '$0.00' );
        $worksheet->write( 3, 0, 49.99, $format04 );        # $49.99

      	#ע�⣬��Ҳ����ʹ����������Ӣ������Ԫ�Ļ��ҷ��š�
        #�������ҿ���Ҫ��ʹ��Unicode��
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

        # ��������
        $format13->set_num_format( '00000' );
        $worksheet->write( 14, 0, '01209', $format13 );

  
	��ɫ�ĸ�ʽӦ��ʹ������ֵ֮һ��

        [Black] [Blue] [Cyan] [Green] [Magenta] [Red] [White] [Yellow]

    Alternatively you can specify the colour based on a colour index as
    follows: "[Color n]", where n is a standard Excel colour index - 7. See
    the 'Standard colors' worksheet created by formats.pl.
	��Ϊѡ���������Ը�����������ɫ����ָ����ɫ�� "[Color n]"��n�Ǳ�׼��Excel��ɫ����-7.�鿴��formats.pl�����Ĺ������е�'Standard colors'��

   
    You should ensure that the format string is valid in Excel prior to
    using it in WriteExcel.
	��Ӧ��ȷ����ʽ�ַ����� WriteExcel��ʹ����֮ǰ�ǺϷ���.

    �����ı���ʽ��Excel�ڽ��ĸ�ʽ��
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

	�鿴 formats.pl.�е�'Numerical formats'��������
	���ɲ鿴number_formats1.html ��number_formats2.html 

    Note 1. Numeric formats 23 to 36 are not documented by Microsoft and may
    differ in international versions.
	ע�⣬23��36֮�������ָ�ʽ��Microsoft��û��˵���ĵ������ڲ�֮���İ汾���ܲ�ͬ��

    Note 2. In Excel 5 the dollar sign appears as a dollar sign. In Excel
    97-2000 it appears as the defined local currency symbol.
	ע��2����Excel5����Ԫ��������Ԫ���ų��֡���Excel97-2000����Ϊ���ض����Ļ��ҷ��ų��֡�

  set_locked()
        Default state:      Cell locking is on
        Default action:     Turn locking on
        Valid args:         0, 1

    This property can be used to prevent modification of a cells contents.
    Following Excel's convention, cell locking is turned on by default.
    However, it only has an effect if the worksheet has been protected, see
    the worksheet "protect()" method.
	���������ڷ�ֹ�޸ĵ�Ԫ�������ݡ�����Excel��Լ������Ԫ��Ĭ�ϱ�������Ȼ����ֻ�е�������������ʱ�����������á��鿴"protect()"������

        my $locked = $workbook->add_format();
        $locked->set_locked( 1 );    # A non-op

        my $unlocked = $workbook->add_format();
        $locked->set_locked( 0 );

        # ��������������
        $worksheet->protect();

        # �õ�Ԫ�����ܱ��༭.
        $worksheet->write( 'A1', '=1+2', $locked );

        # ������Ԫ���ܱ��༭.
        $worksheet->write( 'A2', '=1+2', $unlocked );

  
	ע�⣺��ʹ��������Ҳ���ṩ�˺����ı������鿴��"protect()" �����йص�ע�����

  set_hidden()
        Default state:      Formula hiding is off
        Default action:     Turn hiding on
        Valid args:         0, 1

    This property is used to hide a formula while still displaying its
    result. This is generally used to hide complex calculations from end
    users who are only interested in the result. It only has an effect if
    the worksheet has been protected, see the worksheet "protect()" method.
	��������������һ����ʽ����Ȼ��ʾ�ù�ʽ�Ľ�������ͨ�����ڶ�ֻ���Ľ������ն��û����ظ��ӵļ������̡�ֻ�е���������������ʱ���÷����������ã��鿴"protect()" ������

        my $hidden = $workbook->add_format();
        $hidden->set_hidden();

        # ��������������
        $worksheet->protect();

        # ��������Ԫ���еĹ�ʽ���ɼ�
        $worksheet->write( 'A1', '=1+2', $hidden );

   ע�⣺��ʹʹ�����룬��Ҳ�����ṩ�˺����ı������鿴����"protect()"������ע�����

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
	�÷��������ڵ�Ԫ���������ı���ˮƽ�ʹ�ֱ���뷽ʽ����ֱ��ˮƽ���뷽ʽ���Խ��ϡ��÷����÷����£�

        my $format = $workbook->add_format();
        $format->set_align( 'center' );
        $format->set_align( 'vcenter' );
        $worksheet->set_row( 0, 30 );
        $worksheet->write( 0, 0, 'X', $format );

    Text can be aligned across two or more adjacent cells using the
    "center_across" property. However, for genuine merged cells it is better
    to use the "merge_range()" worksheet method.
	ʹ��"center_across"���ԣ��ı�������2�����������ڵĵ�Ԫ��֮�����롣Ȼ����������ʵ�ĺϲ���Ԫ������ʹ��"merge_range()"������������

    The "vjustify" (vertical justify) option can be used to provide
    automatic text wrapping in a cell. The height of the cell will be
    adjusted to accommodate the wrapped text. To specify where the text
    wraps use the "set_text_wrap()" method.
	�ڵ�Ԫ���У�"vjustify"����ֱ������ѡ���������ṩ�Զ��ı����ơ���Ԫ���ĸ߶Ȼᱻ�Զ���������Ӧ�����ı���ʹ��"set_text_wrap()"����ָ���ı����Ƶ�λ�á�

   	�鿴formats.pl���ɵ�'Alignment' ��������ȡ����ʵ����

  set_center_across()
        Default state:      Center across selection is off
        Default action:     Turn center across on
        Valid args:         1

    Text can be aligned across two or more adjacent cells using the
    "set_center_across()" method. This is an alias for the
    "set_align('center_across')" method call.
	ʹ��"set_center_across()"�������ı��ܹ���2�����������ڵ�Ԫ��֮�����롣���� "set_align('center_across')" �������õı�����

    Only one cell should contain the text, the other cells should be blank:
	Ӧ��ֻ��һ����Ԫ�������ı���������Ԫ���ǿյģ�

        my $format = $workbook->add_format();
        $format->set_center_across();

        $worksheet->write( 1, 1, 'Center across selection', $format );
        $worksheet->write_blank( 1, 2, $format );

    See also the "merge1.pl" to "merge6.pl" programs in the "examples"
    directory and the "merge_range()" method.
	�鿴"merge1.pl" ��"merge6.pl" ������"merge_range()"������

  set_text_wrap()
        Default state:      Text wrap is off
        Default action:     Turn text wrap on
        Valid args:         0, 1

    �����и����� using the text wrap property, the escape character
    "\n" is used to indicate the end of line:
	�����и�ʹ���ı��������Ե����ӣ�

        my $format = $workbook->add_format();
        $format->set_text_wrap();
        $worksheet->write( 0, 0, "It's\na bum\nwrap", $format );

    Excel will adjust the height of the row to accommodate the wrapped text.
    A similar effect can be obtained without newlines using the
    "set_align('vjustify')" method. See the "textwrap.pl" program in the
    "examples" directory.
	Excel�������и�����Ӧ�����ı���ʹ��"set_align('vjustify')" ���������о��ܻ������Ƶ�Ч�����鿴"textwrap.pl"������

  set_rotation()
        Default state:      Text rotation is off
        Default action:     None
        Valid args:         Integers in the range -90 to 90 and 270

    Set the rotation of the text in a cell. The rotation can be any angle in
    the range -90 to 90 degrees.
	��Excel�������ı���ת����ת�Ķ���������-90�ȵ�90��֮�䡣

        my $format = $workbook->add_format();
        $format->set_rotation( 30 );
        $worksheet->write( 0, 0, 'This text is rotated', $format );

    The angle 270 is also supported. This indicates text where the letters
    run from top to bottom.
	Ҳ֧��270����ת���������ı��е���ĸ�Ӷ�����ת���ײ�����

  set_indent()
        Default state:      Text indentation is off
        Default action:     Indent text 1 level
        Valid args:         Positive integers

    This method can be used to indent text. The argument, which should be an
    integer, is taken as the level of indentation:
	�÷������������ı�������Ӧ����һ����������Ϊ�����ļ�����

        my $format = $workbook->add_format();
        $format->set_indent( 2 );
        $worksheet->write( 0, 0, 'This text is indented', $format );

    Indentation is a horizontal alignment property. It will override any
    other horizontal properties but it can be used in conjunction with
    vertical properties.
	������ˮƽ�������ԡ����Ḳ�������κ�ˮƽ���Ե������봹ֱ����һ��ʹ�á�

  set_shrink()
        Default state:      Text shrinking is off
        Default action:     Turn "shrink to fit" on
        Valid args:         1

    This method can be used to shrink text so that it fits in a cell.
	�÷������������ı�����Ӧ��Ԫ���Ĵ�С��

        my $format = $workbook->add_format();
        $format->set_shrink();
        $worksheet->write( 0, 0, 'Honey, I shrunk the text!', $format );

  set_text_justlast()
        Default state:      Justify last is off
        Default action:     Turn justify last on
        Valid args:         0, 1

    Only applies to Far Eastern versions of Excel.
	ֻ��Զ���汾��Excel���á�

  set_pattern()
        Ĭ��״̬:      Pattern is off
        Ĭ����Ϊ:      Solid fill is on
        �Ϸ�����:      0 .. 18

    Set the background pattern of a cell.
	���õ�Ԫ���ı���ͼ����

    Examples of the available patterns are shown in the 'Patterns' worksheet
    created by formats.pl. However, it is unlikely that you will ever need
    anything other than Pattern 1 which is a solid fill of the background
    color.
	����ͼ����������ʾ��formats.pl������'Patterns'�������С�Ȼ��������ͼ��1�Ǳ���ɫ����ȫ�������㲻����Ҫ���������ǲ����ܵġ���

  set_bg_color()
        Default state:      Color is off
        Default action:     Solid fill.
        Valid args:         See set_color()

    The "set_bg_color()" method can be used to set the background colour of
    a pattern. Patterns are defined via the "set_pattern()" method. If a
    pattern hasn't been defined then a solid fill pattern is used as the
    default.
	"set_bg_color()"������������ͼ���ı�����ɫ��ͼ��ͨ��"set_pattern()"�������塣����û�ж���ͼ������Ĭ��ʹ����ȫ����ͼ����

    �����и������ڵ�Ԫ����������ȫ���������ӣ�
	of how to set up a solid fill in a cell:

        my $format = $workbook->add_format();

        $format->set_pattern();    # ʹ����ȫ���������ǿ�ѡ��

        $format->set_bg_color( 'green' );
        $worksheet->write( 'A1', 'Ray', $format );

	�鿴formats.pl�����е�'Patterns'��������ȡ�������ӡ�

  set_fg_color()
        Default state:      Color is off
        Default action:     Solid fill.
        Valid args:         See set_color()

    "set_fg_color()"������������ͼ����ǰ��ɫ��
    �鿴formats.pl�����е�'Patterns'��������ȡ�������ӡ�

  set_border()
        Also applies to:    set_bottom()
                            set_top()
                            set_left()
                            set_right()

        Default state:      Border is off
        Default action:     Set border type 1
        Valid args:         0-13, See below.

	��Ԫ���߿��ɵײ��ġ������ġ������ġ��Ҳ��ı߿����ɡ���Щ�߿���ʹ��"set_border()"����Ϊͬ������ɫ���򵥶�ʹ������չʾ�����ط������á�

  
	������ʾ����Excel::Writer::XLSX���������������ı߿���ʽ��

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

	������ʾ�˰���ʽ�������ı߿���

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

	������ʽ����Excel�Ի������������ı߿���

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
	���õı߿���ʽ��������ʾ����formats.pl������'Borders'�������С�

  set_border_color()
        Also applies to:    set_bottom_color()
                            set_top_color()
                            set_left_color()
                            set_right_color()

        Default state:      Color is off
        Default action:     Undefined
        Valid args:         See set_color()


	���õ�Ԫ���߿�����ɫ����Ԫ���߿��ɵױ߿򡢶��߿������߿����ұ߿����ɡ�
	��Щ�߿���ʹ��"set_border()"����Ϊͬ������ɫ���򵥶�ʹ������չʾ�����ط������á�
	�߿���ʽ����ɫ��������ʾ����formats.pl���򴴽��� 'Borders'�������С�

  copy( $format )
    This method is used to copy all of the properties from one Format object
    to another:
	�÷������ڴ�һ����ʽ�����и������е����Ե���һ����ʽ�����У�

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
	"copy()"����ֻ������ʹ�ø÷����ӿڵĸ�ʽ���������õġ�������ֱ��ʹ��ɢ�����ø�ʽ�����ԣ���һ�㲻��Ҫcopy()������

 
	ע�⣺�ⲻ��һ�����ƹ��캯�����ڸ���֮ǰ2�����󶼱����Ǵ��ڵġ�

UNICODE IN EXCEL

	�������� "Excel::Writer::XLSX"�д���Unicode�ļ��顣

   
	Excel::Writer::XLSX��Spreadsheet::WriteExcel ��д�뷽ʽ��ͬ������ֻ����UTF-8��ʽ��Unicode���ݣ����Ҳ��ܴ���������UTF-16��Excel��ʽ��

	����������UTF-8��ʽ�ģ��� Excel::Writer::XLSX ���Զ���������

    �����㴦�����Ƿ�UTF-8��ʽ��non-ASCII�ַ�����perl���ṩ���õ�Encode����ģ��������ת��Ϊ��Ҫ�ĸ�ʽ�����磺

        use Encode 'decode';

        my $string = 'some string with koi8-r characters';
           $string = decode('koi8-r', $string); # koi8-r to utf8

    Alternatively you can read data from an encoded file and convert it to
    "UTF-8" as you read it in:
	��Ϊѡ�񣬵�����������ʱ�����ܴ�һ���������ļ��ж�ȡ���ݲ�������ת��ΪUTF-8��

        my $file = 'unicode_koi8r.txt';
        open FH, '<:encoding(koi8-r)', $file or die "Couldn't open $file: $!\n";

        my $row = 0;
        while ( <FH> ) {
            # Data read in is now in utf8 format.
            chomp;
            $worksheet->write( $row++, 0, $_ );
        }

    Ҳ���鿴"unicode_*.pl"������

COLOURS IN EXCEL
    Excel provides a colour palette of 56 colours. In Excel::Writer::XLSX
    these colours are accessed via their palette index in the range 8..63.
    This index is used to set the colour of fonts, cell patterns and cell
    borders. ���磬:
	Excel�ṩ��56����ɫ�ĵ�ɫ�塣��Excel::Writer::XLSX����Щ��ɫͨ�����ǵ�������ֵ�������ʣ�����ֵ��Χ��8..63���˴�������ֵ��������������ɫ����Ԫ��ͼ���͵�Ԫ���߿������磺

        my $format = $workbook->add_format(
                                            color => 12, # index for blue
                                            font  => 'Arial',
                                            size  => 12,
                                            bold  => 1,
                                         );


	��õ���ɫҲ��ͨ�����ǵ����ַ��ʡ�������Ϊ��ɫ�����ļ򵥱�����

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

    ����:

        my $font = $workbook->add_format( color => 'red' );

 	Excel��VBA�û�Ӧ��ע���ȼ۵���ɫ������1..56����8..63.

    If the default palette does not provide a required colour you can
    override one of the built-in values. This is achieved by using the
    "set_custom_color()" workbook method to adjust the RGB (red green blue)
    components of the colour:
	����Ĭ�ϵ����ϲ����ṩ����Ҫ����ɫ����������д���е�����ֵ��ʹ��"set_custom_color()"������������������ɫ��RGB���� �� �����ɷֿ����������㣺
        my $ferrari = $workbook->set_custom_color( 40, 216, 12, 12 );

        my $format = $workbook->add_format(
            bg_color => $ferrari,
            pattern  => 1,
            border   => 1
        );

        $worksheet->write_blank( 'A1', $format );

    �鿴"colors.pl" ������
DATES AND TIME IN EXCEL
    
	����Excel�е����ں�ʱ�䣬��2����Ҫ�����飺

    1��Excel�е�����/ʱ�� ��һ��ʵ������һ��Excel���ָ�ʽ��
   	2��Excel::Writer::XLSX �е�"write()"���������Զ���������/ʱ�䡱�ַ���ת����Excel�ġ�����/ʱ�䡯��
    ���������Ĺ���������ʱ��������ת������Ҫ�ĸ�ʽ��һЩ���飬��2�����и���ϸ�Ľ��͡�

	Excel�ġ�����/ʱ�䡱�������ּ��ϸ�ʽ
    If you write a date string with "write()" then all you will get is a
    string:
	������ʹ��"write()"����д�������ַ������������㽫�õ��Ļ���һ���ַ�����

        $worksheet->write( 'A1', '02/03/04' );   # !! д��һ���ַ�������һ������. !!

 	Excel�����ں����ִ���ʵ�������磬"Jan 1 2001 12:30 AM"��������36892.521.

	���ݵ��������ִ洢�����Լ�Ԫ������������С�����ִ洢����һ���İٷֱȡ�

    Excel�е����ڻ�ʱ���������κ��������ơ�Ϊ�������������ڵ���ʽ��ʾ�������뽫һ��Excel���ָ�ʽӦ�õ����������ϡ�������һЩ���ӣ�

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

  Excel::Writer::XLSX ���Զ�ת��������/ʱ�䡱�ַ���
    Excel::Writer::XLSX doesn't automatically convert input date strings
    into Excel's formatted date numbers due to the large number of possible
    date formats and also due to the possibility of misinterpretation.
	���ڿ��õ����ڸ�ʽ�����ܴ���Ҳ���ڿ��ܵ����⣬Excel::Writer::XLSX���ܽ������������ַ����Զ�ת��ΪExcel�ĸ�ʽ���������֡�

    ���磬, does "02/03/04" mean March 2 2004, February 3 2004 or even
    March 4 2002.
	���磬"02/03/04"����˼�� March 2 2004, February 3 2004 ������ March 4 2002����

    Therefore, in order to handle dates you will have to convert them to
    numbers and apply an Excel format. Some methods for converting dates are
    listed in the next section.
	���ˣ�Ϊ�˴������������뽫����ת��Ϊ���ֲ�Ӧ��һ��Excel��ʽ��ת�����ڵ�һЩ�������������½����г���


    ��ֱ�ӵķ�ʽ�ǽ���������ת��ΪISO8601"yyyy-mm-ddThh:mm:ss.sss" ���ڸ�ʽ����ʹ��"write_date_time()"����������:
        $worksheet->write_date_time( 'A2', '2001-01-01T12:20', $format );

    �鿴�ĵ���"write_date_time()"�½ڻ�ȡ��ϸ��Ϣ��

  
	���������ַ���һ�㷽����ʹ��"write_date_time()"��

        1.ʹ����������ʽʶ������������/ʱ�䡣
		2.ʹ��ͬ������������ʽ��ȡ����/ʱ�������ɲ��֡�
		3.������/ʱ��ת��ΪISO8601��ʽ��
        4.ʹ�� write_date_time()�����ָ�ʽд������/ʱ�䡣

    �����и�����:

        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        my $workbook  = Excel::Writer::XLSX->new( 'example.xlsx' );
        my $worksheet = $workbook->add_worksheet();

        # Ϊ��������Ĭ�ϸ�ʽ
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


	���߼��ķ���������ͨ��"add_write_handler()"�����޸�"write()"������������ѡ�������ݸ�ʽ���鿴"add_write_handler()"�½ں�write_handler3.pl�� write_handler4.pl������

  Converting dates and times to an Excel date or time
  �����ں�ʱ��ת��ΪExcel�����ڻ�ʱ��

	������"write_date_time()" ����ֻ�Ǵ������ں�ʱ���ķ���֮һ��

    You can also use the "convert_date_time()" worksheet method to convert
    from an ISO8601 style date string to an Excel date and time number.
	��Ҳ����ʹ��"convert_date_time()"������������ISO8601�������ַ���ת��ΪExcel�����ں�ʱ�����֡�

     Excel::Writer::XLSX::Utilityģ��������ʱ�䴦��������

        use Excel::Writer::XLSX::Utility;

        $date           = xl_date_list(2002, 1, 1);         # 37257
        $date           = xl_parse_date("11 July 1997");    # 35622
        $time           = xl_parse_time('3:21:36 PM');      # 0.64
        $date           = xl_decode_date_EU("13 May 2002"); # 37389

    ע�⣬��Щ������Ҫ������CPANģ�顣

    
OUTLINES AND GROUPING IN EXCEL   Excel�е����ͷּ���ʾ
    Excel allows you to group rows or columns so that they can be hidden or
    displayed with a single mouse click. This feature is referred to as
    outlines.
	Excel�����㽫�л��з��飬��ʹ��������ʱ�����ܱ����ػ���ʾ���ù��ܽ����ּ���ʾ��

    Outlines can reduce complex data down to a few salient sub-totals or
    summaries.
	�ּ���ʾ�ܽ��������ݼ��ٵ�����ͻ����С�ƻ��ܽᡣ

    This feature is best viewed in Excel but the following is an ASCII
    representation of what a worksheet with three outlines might look like.
    Rows 3-4 and rows 7-8 are grouped at level 2. Rows 2-9 are grouped at
    level 1. The lines at the left hand side are called outline level bars.
	�ù���������Excel�в鿴�����������Ǵ���3���ּ���ʾ�Ĺ�������������ASCIIͼ��
	3-4����7-8���ڵ�2�������顣2-9���ڵ�һ���𱻷��顣�������߽����ּ���ʾ����

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
	��ÿ������2��������(-),�ּ����۵�����������һͼ���е����ݡ������ű�Ϊ�Ӻ�ʱ�������ּ���ʾ�е����ݱ����ء�

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
	��������1�еļ��ţ��ּ��ᰴ���·�ʽ�۵�ʣ�����У�

                ------------------------------------------
         1 2 3 |   |   A   |   B   |   C   |   D   |  ...
                ------------------------------------------
               | 1 |   A   |       |       |       |  ...
         +     | . |  ...  |  ...  |  ...  |  ...  |  ...

    Grouping in "Excel::Writer::XLSX" is achieved by setting the outline
    level via the "set_row()" and "set_column()" worksheet methods:
	ͨ��"set_row()" �� "set_column()"�������÷ּ���ʾ��"Excel::Writer::XLSX" ���������ݷ��飺

        set_row( $row, $height, $format, $hidden, $level, $collapsed )
        set_column( $first_col, $last_col, $width, $format, $hidden, $level, $collapsed )

    The following example sets an outline level of 1 for rows 1 and 2
    (zero-indexed) and columns B to G. The parameters $height and $XF are
    assigned default values since they are undefined:
	����������Ϊ1-2�У���0��ʼ��������B-G�������˼���Ϊ1�ķּ���ʾ������$height �� $XFָ����Ĭ��ֵ����Ϊ������δ�����ģ�undefined����

        $worksheet->set_row( 1, undef, undef, 0, 1 );
        $worksheet->set_row( 2, undef, undef, 0, 1 );
        $worksheet->set_column( 'B:G', undef, undef, 0, 1 );

    Excel allows up to 7 outline levels. Therefore the $level parameter
    should be in the range "0 <= $level <= 7".
	Excel��������7���ּ���ʾ�����ˣ�$level����Ӧ���ڷ�Χ "0 <= $level <= 7"��

    Rows and columns can be collapsed by setting the $hidden flag for the
    hidden rows/columns and setting the $collapsed flag for the row/column
    that has the collapsed "+" symbol:
	ͨ��Ϊ���ص��л�������$hidden���ǲ���Ϊ����"+"�ŵ��л�������$collapsed�������۵��л��У�

        $worksheet->set_row( 1, undef, undef, 1, 1 );
        $worksheet->set_row( 2, undef, undef, 1, 1 );
        $worksheet->set_row( 3, undef, undef, 0, 0, 1 );          # Collapsed flag.

        $worksheet->set_column( 'B:G', undef, undef, 1, 1 );
        $worksheet->set_column( 'H:H', undef, undef, 0, 0, 1 );   # Collapsed flag.

    Note: Setting the $collapsed flag is particularly important for
    compatibility with OpenOffice.org and Gnumeric.
	ע�⣺����$collapsed���Ƕ��ڼ���OpenOffice.org �͵��ӱ����ر���Ҫ��

    �鿴"outline.pl"��"outline_collapsed.pl" ������

    Some additional outline properties can be set via the
    "outline_settings()" worksheet method, see above.
	һЩ�����ķּ���ʾ������ͨ��"outline_settings()"�������������ã��鿴���������ӡ�

DATA VALIDATION IN EXCEL Excel�е�������֤
    Data validation is a feature of Excel which allows you to restrict the
    data that a users enters in a cell and to display help and warning
    messages. It also allows you to restrict input to values in a drop down
    list.
	������֤��Excel��һ�ֹ��ܣ��������������û��ڵ�Ԫ�������������ݲ�����ʾ�����;�����Ϣ����Ҳ��������һ�������б�����������ֵ��

    A typical use case might be to restrict data in a cell to integer values
    in a certain range, to provide a help message to indicate the required
    value and to issue a warning if the input data doesn't meet the stated
    criteria. In Excel::Writer::XLSX we could do that as follows:
	һ�����͵�ʹ��ʵ����������һ����Χ�ڽ���Ԫ���е���������Ϊ�������������������ݲ����ϱ�׼�������ṩ������Ϣָ����Ҫ��ֵ�򷢳�һ�����档��Excel::Writer::XLSX�����ǿ���ʹ�����µķ�����

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
	�������½�����������ʹ��"data_validation()"���������ĸ���ѡ�

  data_validation( $row, $col, { parameter => 'value', ... } )
    The "data_validation()" method is used to construct an Excel data
    validation.
	"data_validation()"�������ڹ���һ��Excel������֤��

    It can be applied to a single cell or a range of cells. You can pass 3
    parameters such as "($row, $col, {...})" or 5 parameters such as
    "($first_row, $first_col, $last_row, $last_col, {...})". You can also
    use "A1" style notation. ���磬:
	�������ڵ�����Ԫ����һ����Χ�ڵĵ�Ԫ���������Դ���3����������"($row, $col, {...})"��5���������� "($first_row, $first_col, $last_row, $last_col, {...})"����Ҳ����ʹ��A1�����ı�ʾ�������磺

        $worksheet->data_validation( 0, 0,       {...} );
        $worksheet->data_validation( 0, 0, 4, 1, {...} );

        # Which are the same as:

        $worksheet->data_validation( 'A1',       {...} );
        $worksheet->data_validation( 'A1:B5',    {...} );

     
    The last parameter in "data_validation()" must be a hash ref containing
    the parameters that describe the type and style of the data validation.
    The allowable parameters are:
	"data_validation()"�е�����һ������������һ����������������֤�����ͺͷ���������ɢ�����á������Ĳ����ǣ�
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
	��Щ�������������½����������������������ǿ�ѡ�ģ�Ȼ������ͨ����Ҫ������Ҫѡ��"validate", "criteria" �� "value".

        $worksheet->data_validation('B3',
            {
                validate => 'integer',
                criteria => '>',
                value    => 100,
            });

     "data_validation" ��������:

         0 �ɹ�.
        -1 ������������.
        -2 �л��г���.
        -3 ������ֵ����ȷ.

  validate
    This parameter is passed in a hash ref to "data_validation()".
	�˲�����ɢ�������б����ݸ�"data_validation()"��

	"validate"������������������֤���������͡��ò���������Ҫ�Ĳ���û��Ĭ��ֵ��
	������ֵ�ǣ�

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
		any����ָ�����������������Ƶġ����벻ʹ��������֤��ͬ����ֻΪ�������ṩ���Ҳ�����Excel::Writer::XLSX�����о���ʹ�á�

    *   integer restricts the cell to integer values. Excel refers to this
        as 'whole number'.
		integer���Ƶ�Ԫ����ֵΪ������Excel��������Ϊ������

            validate => 'integer',
            criteria => '>',
            value    => 100,

    *   decimal���Ƶ�Ԫ����ֵΪʮ����ֵ��

            validate => 'decimal',
            criteria => '>',
            value    => 38.6,

    *   list restricts the cell to a set of user specified values. These can
        be passed in an array ref or as a cell range (named ranges aren't
        currently supported):
		list���Ƶ�Ԫ����ֵΪһ���û�ָ����ֵ����Щֵ�����������û���Ԫ����Χ��Ŀǰ��֧��������Χ���д��ݣ�

            validate => 'list',
            value    => ['open', 'high', 'close'],
            # Or like this:
            value    => 'B1:B3',

        Excel requires that range references are only to cells on the same
        worksheet.
		ExcelҪ��ֵ������ֻ������ͬһ�������ĵ�Ԫ���ġ�

    *   date restricts the cell to date values. Dates in Excel are expressed
        as integer values but you can also pass an ISO860 style string as
        used in "write_date_time()". See also "DATES AND TIME IN EXCEL" for
        more information about working with Excel's dates.
		date���Ƶ�Ԫ����ֵΪ���ڡ�Excel�е����ڱ�����Ϊ����������Ҳ������"write_date_time()"ʹ�õ�����������һ��ISO860�������ַ�����
            validate => 'date',
            criteria => '>',
            value    => 39653, # 24 July 2008
            # Or like this:
            value    => '2008-07-24T',

    *   time restricts the cell to time values. Times in Excel are expressed
        as decimal values but you can also pass an ISO860 style string as
        used in "write_date_time()". See also "DATES AND TIME IN EXCEL" for
        more information about working with Excel's times.
		time���Ƶ�Ԫ����ֵΪʱ�䡣Excel�е�ʱ�䱻����Ϊʮ����ֵ������Ҳ������"write_date_time()"ʹ�õ�����������һ��ISO860�������ַ�����

            validate => 'time',
            criteria => '>',
            value    => 0.5, # Noon
            # Or like this:
            value    => 'T12:00:00',

    *   length restricts the cell data based on an integer string length.
        Excel refers to this as 'Text length'.
		length����һ�������ַ����������Ƶ�Ԫ�����ݡ�Excel����ֵ����Ϊ�ı����ȡ�

            validate => 'length',
            criteria => '>',
            value    => 10,

    *   custom restricts the cell based on an external Excel formula that
        returns a "TRUE/FALSE" value.
		custom���ݷ��ء�TRUE/FALSE��ֵ���ⲿExcel��ʽ���Ƶ�Ԫ����
            validate => 'custom',
            value    => '=IF(A10>B10,TRUE,FALSE)',

  criteria
    This parameter is passed in a hash ref to "data_validation()".
    �ò�����һ��ɢ�������д��ݵ�"data_validation()"��
    The "criteria" parameter is used to set the criteria by which the data
    in the cell is validated. It is almost always required except for the
    "list" and "custom" validate options. It has no default value. Allowable
    values are:
	"criteria"�����������õ�Ԫ������֤�����������õı�׼��������������Ҫ������ "list" �� "custom"��֤ѡ���û��Ĭ��ֵ��������ֵΪ��

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
	��Ҳ����ʹ��Excel���ı������ַ������������ĵ�һ���У�������ͨ�ı�ѡ���š��������ǵȼ۵ģ�

        validate => 'integer',
        criteria => 'greater than',
        value    => 100,

        validate => 'integer',
        criteria => '>',
        value    => 100,

    The "list" and "custom" validate options don't require a "criteria". If
    you specify one it will be ignored.
	"list" �� "custom"��Ч��ѡ���Ҫ�Ը���׼��������ָ��һ�����ᱻ���ԡ�

        validate => 'list',
        value    => ['open', 'high', 'close'],

        validate => 'custom',
        value    => '=IF(A10>B10,TRUE,FALSE)',

  value | minimum | source
    This parameter is passed in a hash ref to "data_validation()".
    �ò�����ɢ�������б����ݸ� "data_validation()"��
    The "value" parameter is used to set the limiting value to which the
    "criteria" is applied. It is always required and it has no default
    value. You can also use the synonyms "minimum" or "source" to make the
    validation a little clearer and closer to Excel's description of the
    parameter:
	"value"�������ڶ�Ӧ����"criteria"��ֵ���ü���ֵ�������Ǳ���Ҫ��������û��Ĭ��ֵ����Ҳ����ʹ��ͬ���� "minimum"��"source"����Ч�Լ���������������Excel�Ĳ���������������

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
    �ò�����ɢ�������б����ݸ�"data_validation()"��
    The "maximum" parameter is used to set the upper limiting value when the
    "criteria" is either 'between' or 'not between':
	
	��"criteria"��ֵ�� 'between' �� 'not between'ʱ��"maximum"������������ֵ�����ޡ�

        validate => 'integer',
        criteria => 'between',
        minimum  => 1,
        maximum  => 100,

  ignore_blank
    This parameter is passed in a hash ref to "data_validation()".
    �ò�����ɢ�������б����ݸ�"data_validation()"��
    The "ignore_blank" parameter is used to toggle on and off the 'Ignore
    blank' option in the Excel data validation dialog. When the option is on
    the data validation is not applied to blank data in the cell. It is on
    by default.
	
	"ignore_blank"����������Excel��������Ч�Լ����Ի����п������ر�'Ignore blank'ѡ�����ѡ���ʱ��������Ч�Լ��鲻��Ӧ�õ���Ԫ���еĿհ������ϡ�Ĭ�����ǿ����ġ�

        ignore_blank => 0,  # Turn the option off

  dropdown
    This parameter is passed in a hash ref to "data_validation()".
    �ò�����ɢ�������б����ݸ�"data_validation()"��
    The "dropdown" parameter is used to toggle on and off the 'In-cell
    dropdown' option in the Excel data validation dialog. When the option is
    on a dropdown list will be shown for "list" validations. It is on by
    default.
	
	"dropdown"����������Excel��������Ч�ԶԻ����п������ر�'In-cell dropdown'ѡ���������ѡ��ʱ������Ϊ�б���֤�����������б���Ĭ�����ǿ����ġ���

        dropdown => 0,      # Turn the option off

  input_title
    This parameter is passed in a hash ref to "data_validation()".
    �ò�����ɢ�������б����ݸ�"data_validation()"��
    The "input_title" parameter is used to set the title of the input
    message that is displayed when a cell is entered. It has no default
    value and is only displayed if the input message is displayed. See the
    "input_message" parameter below.
	"input_title"������������������Ϣ�ı��⣬��û��Ĭ��ֵ������ֻ�е�������Ϣ��ʾʱ�ų��֡��鿴������ "input_message" ������

        input_title   => 'This is the input title',

    The maximum title length is 32 characters.
	�����ı��ⳤ����32���ַ���

  input_message
        �ò�����ɢ�������б����ݸ�"data_validation()"��


    The "input_message" parameter is used to set the input message that is
    displayed when a cell is entered. It has no default value.
	"input_message"�����������ü��뵥Ԫ��ʱ��ʾ��������Ϣ����û��Ĭ��ֵ��

        validate      => 'integer',
        criteria      => 'between',
        minimum       => 1,
        maximum       => 100,
        input_title   => 'Enter the applied discount:',
        input_message => 'between 1 and 100',

    The message can be split over several lines using newlines, "\n" in
    double quoted strings.
	��Ϣ����ʹ�û��зָ�Ϊ���С�"\n"��˫�����ַ����С�

        input_message => "This is\na test.

    The maximum message length is 255 characters.
	��Ϣ�����󳤶���255���ַ���

  show_input
    This parameter is passed in a hash ref to "data_validation()".
    �ò�����ɢ�������б����ݸ�"data_validation()"��
    The "show_input" parameter is used to toggle on and off the 'Show input
    message when cell is selected' option in the Excel data validation
    dialog. When the option is off an input message is not displayed even if
    it has been set using "input_message". It is on by default.
	
	"show_input"����������Excel��������Ч�Լ����Ի����п������ر�'Show input message when cell is selected'ѡ�����ѡ���ر�ʱ��������Ϣ������ʾ����ʹ��������"input_message"��Ĭ�����ǿ����ġ�
        show_input => 0,      # Turn the option off

  error_title
    This parameter is passed in a hash ref to "data_validation()".
	�ò�����ɢ�������б����ݸ�"data_validation()"��

    The "error_title" parameter is used to set the title of the error
    message that is displayed when the data validation criteria is not met.
    The default error title is 'Microsoft Excel'.

        error_title   => 'Input value is not valid',

    The maximum title length is 32 characters.
	���������󳤶���32���ַ���

  error_message
    This parameter is passed in a hash ref to "data_validation()".
     �ò�����ɢ�������б����ݸ�"data_validation()"��
    The "error_message" parameter is used to set the error message that is
    displayed when a cell is entered. The default error message is "The
    value you entered is not valid.A user has restricted values that can
    be entered into the cell.".
	
	"error_message" �����������ü��뵥Ԫ��ʱ��ʾ��������Ϣ��Ĭ�ϵĴ�����Ϣ��"The
    value you entered is not valid."���û������������뵽��Ԫ���е�ֵ��


        validate      => 'integer',
        criteria      => 'between',
        minimum       => 1,
        maximum       => 100,
        error_title   => 'Input value is not valid',
        error_message => 'It should be an integer between 1 and 100',

    The message can be split over several lines using newlines, "\n" in
    double quoted strings.
	��Ϣ����ʹ�û��зָ�Ϊ���С�"\n"��˫�����ַ����С�
	
	

        input_message => "This is\na test.",

    The maximum message length is 255 characters.
	��������Ϣ����ֵ��255���ַ���

  error_type
    This parameter is passed in a hash ref to "data_validation()".
	�ò�����ɢ�������б����ݸ�"data_validation()"��

    The "error_type" parameter is used to specify the type of error dialog
    that is displayed. There are 3 options:
	"error_type"��������ָ�����ֵĴ����Ի��������͡���3��ѡ�

        'stop'
        'warning'
        'information'

    Ĭ����'stop'.

  show_error
    �ò�����ɢ�������б����ݸ�"data_validation()".

    The "show_error" parameter is used to toggle on and off the 'Show error
    alert after invalid data is entered' option in the Excel data validation
    dialog. When the option is off an error message is not displayed even if
    it has been set using "error_message". It is on by default.
	
	"show_error"����������Excel��������Ч�Լ����Ի����п������ر�'Show error
    alert after invalid data is entered'ѡ�����ѡ���ر�ʱ��������Ϣ������ʾ����ʹ��������"error_message"��Ĭ�����ǿ����ġ�

        show_error => 0,      # Turn the option off

  Data Validation Examples
    Example 1. Limiting input to an integer greater than a fixed value.
	��1.�������޶�Ϊ��ĳһ�̶�ֵ����������

        $worksheet->data_validation('A1',
            {
                validate        => 'integer',
                criteria        => '>',
                value           => 0,
            });

    Example 2. Limiting input to an integer greater than a fixed value where
    the value is referenced from a cell.
   ��2.�������޶�Ϊ��ĳһ�̶�ֵ�����������ù̶�ֵ���Ե�Ԫ�����á�
        $worksheet->data_validation('A2',
            {
                validate        => 'integer',
                criteria        => '>',
                value           => '=E3',
            });

    Example 3. Limiting input to a decimal in a fixed range.
	��3.����������Ϊĳһ�̶���Χ�ڵ�ʮ����ֵ��

        $worksheet->data_validation('A3',
            {
                validate        => 'decimal',
                criteria        => 'between',
                minimum         => 0.1,
                maximum         => 0.5,
            });

    Example 4. Limiting input to a value in a dropdown list.
	��4. ����������Ϊ�����б��е�ĳ��ֵ��

        $worksheet->data_validation('A4',
            {
                validate        => 'list',
                source          => ['open', 'high', 'close'],
            });

    Example 5. Limiting input to a value in a dropdown list where the list
    is specified as a cell range.
	��5.����������Ϊ�����б��е�ĳ��ֵ���������б��ɵ�Ԫ����Χָ����

        $worksheet->data_validation('A5',
            {
                validate        => 'list',
                source          => '=$E$4:$G$4',
            });

    Example 6. Limiting input to a date in a fixed range.
	��6.����������Ϊĳһ�̶���Χ�ڵ�����ֵ��

        $worksheet->data_validation('A6',
            {
                validate        => 'date',
                criteria        => 'between',
                minimum         => '2008-01-01T',
                maximum         => '2008-12-12T',
            });

    Example 7. Displaying a message when the cell is selected.
	��7.��ѡ�е�Ԫ��ʱ����ʾ��ʾ��Ϣ��

        $worksheet->data_validation('A7',
            {
                validate      => 'integer',
                criteria      => 'between',
                minimum       => 1,
                maximum       => 100,
                input_title   => 'Enter an integer:',
                input_message => 'between 1 and 100',
            });

    �鿴 "data_validate.pl"������

 EXCEL �е�������ʽ
    ������ʽ��Excel��һ��ܣ�����������һ���ı�׼��һ����ʽӦ�õ�һ����Ԫ����һ����Χ�ڵĵ�Ԫ���С�

	���磬�����ı�׼������"conditional_format.pl"������ʹ�ú�ɫ����ֵ���ڻ�����50�ĵ�Ԫ����

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
	"conditional_format()"�������ڸ����û������ı�׼����ʽӦ�õ�Excel::Writer::XLSX�ļ��С�

	���ܱ�Ӧ�õ�������Ԫ���л�һ����Χ�ڵĵ�Ԫ���С������Դ���3�����������磺"($row, $col, {...})" ��5������������ "($first_row, $first_col, $last_row, $last_col, {...})".��Ҳ����ʹ��A1��ʾ�������磺

        $worksheet->conditional_format( 0, 0,       {...} );
        $worksheet->conditional_format( 0, 0, 4, 1, {...} );

        # Which are the same as:

        $worksheet->conditional_format( 'A1',       {...} );
        $worksheet->conditional_format( 'A1:B5',    {...} );

     
     "conditional_format()" ��������һ������������һ��ɢ�����ã��������������ݺϷ��Ե����ͺͷ�������Ҫ������:
	
	"conditional_format()" �����е�����һ������������һ��ɢ�����ã������ð���������������Ч�Ե����ͺ���ʽ�Ĳ�������Ҫ�Ĳ����У�

        type
        format
        criteria  #��׼
        value
        minimum
        maximum

	����ָ��������ʽ���͵����������������������½�����ʾ��

  type
	�ò�����ɢ�������б����ݸ�"conditional_format()"

	"type"����������������Ӧ�õ�������ʽ����������Ҫ�ģ�����û��Ĭ��ֵ��������"type"ֵ�����ǵ��йز����ǣ�

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

	���еĸ�ʽ���Ͷ���"format"�����������档�������ͺͲ�������ͼ�����û��ں��ʵ�ʱ�����ӡ�

  type => 'cell'
    This is the most common conditional formatting type. It is used when a
    format is applied to a cell based on a simple criteria. ���磬:
	������õ�������ʽ���͡�����һ���򵥵ı�׼���ø�ʽ�����ڽ���ʽӦ�õ���Ԫ����ʱ��ʹ�á����磺

        $worksheet->conditional_formatting( 'A1',
            {
                type     => 'cell',
                criteria => 'greater than',
                value    => 5,
                format   => $red_format,
            }
        );

    ����ʹ��"between"��׼:

        $worksheet->conditional_formatting( 'C1:C4',
            {
                type     => 'cell',
                criteria => 'between',
                minimum  => 20,
                maximum  => 30,
                format   => $green_format,
            }
        );

  criteria # ��׼
    The "criteria" parameter is used to set the criteria by which the cell
    data will be evaluated. It has no default value. The most common
    criteria as applied to "{ type => 'cell' }" are:
	"criteria"�����������õ�Ԫ�����ݽ��������ı�׼����û��Ĭ��ֵ���������"{ type => 'cell' }"�ı�׼�У�

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
	��Ҳ����ʹ��Excel�������ַ������������ĵ�һ�У���ʹ�ø���ͨ�ķ��š�

    Additional criteria which are specific to other conditional format types
    are shown in the relevant sections below.
	����ָ��������ʽ���͵�������׼�������������½�����ʾ��

  value
    
	"value"ͨ����"criteria"����һ��ʹ�ã��������ý��������ĵ�Ԫ�����ݵĹ�����

        type     => 'cell',
        criteria => '>',
        value    => 5
        format   => $format,

	"value"����Ҳ�����ǵ�Ԫ�����á�

        type     => 'cell',
        criteria => '>',
        value    => '$C$1',
        format   => $format,

  format
    The "format" parameter is used to specify the format that will be
    applied to the cell when the conditional formatting criteria is set. The
    format is created using the "add_format()" method in the same way as
    cell formats:
	
	��������ʽ��׼���ú���"format"��������ָ������Ӧ�õ���Ԫ���еĸ�ʽ���ø�ʽʹ���뵥Ԫ����ʽһ����"add_format()"����������

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
	������ʽ������Excelͬ���Ĺ����������Ѿ����ڵĵ�Ԫ����ʽ�ص������Ҳ������е������ͱ߿������ܱ��޸ġ������޸ĵ���������������������С���ϱ����±ꡣ���ܱ��޸ĵı߿�������б�߱߿���

    Excel specifies some default formats to be used with conditional
    formatting. You can replicate them using the following
    Excel::Writer::XLSX formats:
	Excelָ����һЩ��������ʽһ��ʹ�õ�Ĭ�ϸ�ʽ��������ʹ��������Excel::Writer::XLSX�ĸ�ʽ��д���ǣ�

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
	����׼��'between' �� 'not between'ʱ��"minimum"������������ֵ�����ޣ�

        validate => 'integer',
        criteria => 'between',
        minimum  => 1,
        maximum  => 100,

  maximum
    The "maximum" parameter is used to set the upper limiting value when the
    "criteria" is either 'between' or 'not between'. See the previous
    example.
	����׼��'between' �� 'not between'ʱ��"maximum"������������ֵ�����ޡ��鿴ǰһ�����ӣ�
  type => 'date'
    The "date" type is the same as "cell" type and uses the same criteria
    and values. However it allows the "value", "minimum" and "maximum"
    properties to be specified in the ISO8601 "yyyy-mm-ddThh:mm:ss.sss" date
    format which is detailed in the "write_date_time()" method.
	"date"������"cell"������ͬ��ʹ����ͬ�ı�׼��ֵ��Ȼ���������� "value", "minimum" �� "maximum"����ָ��ΪISO8601 "yyyy-mm-ddThh:mm:ss.sss"���ڸ�ʽ������"write_date_time()"�����ϸ���ϸ��

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
	"time_period" ��������ָ��Excel��"Dates Occurring"������������ʽ��

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'time_period',
                criteria => 'yesterday',
                format   => $format,
            }
        );

    The period is set in the "criteria" and can have one of the following
    values:
    ������"criteria"�����ã����ҿ��������µ�ֵ֮һ��
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
	
	"text"��������ָ��Excel��"Specific Text"������������ʽ��������ʹ��"criteria" �� "value"�������򵥵��ַ���ƥ�䣺

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'text',
                criteria => 'containing',
                value    => 'foo',
                format   => $format,
            }
        );
 
    "criteria"��ʹ�����µ�ֵ��
        criteria => 'containing',
        criteria => 'not containing',
        criteria => 'begins with',
        criteria => 'ends with',

	"value"����������һ���ַ����򵥸��ַ���

  type => 'average'
    The "average" type is used to specify Excel's "Average" style
    conditional format.
	"average"��������ָ��Excel��"Average"����������ʽ��

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'average',
                criteria => 'above',
                format   => $format,
            }
        );

    The type of average for the conditional format range is specified by the
    "criteria":
	������ʽ��Χ��average������"criteria"ָ������

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
	"duplicate"�������ڸ���һ����Χ����ȫ��ͬ�ĵ�Ԫ����

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'duplicate',
                format   => $format,
            }
        );

  type => 'unique'
    The "unique" type is used to highlight unique cells in a range:
	"unique"�������ڸ���һ����Χ�ڵ�Ψһ��Ԫ����

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'unique',
                format   => $format,
            }
        );

  type => 'top'
    The "top" type is used to specify the top "n" values by number or
    percentage in a range:
	"top"��������ʹ�����ֻ��ٷֱ�ָ����Ԫ����ǰ"n"��ֵ����

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'top',
                value    => 10,
                format   => $format,
            }
        );

    The "criteria" can be used to indicate that a percentage condition is
    required:
	"criteria"�����ڱ�����Ҫһ���ٷֱ�������

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
	"bottom"��������ʹ�����ֻ��ٷֱ�ָ����Ԫ���ĺ���n����ֵ����

    It takes the same parameters as "top", see above.
	���Ĳ����롰top��һ���������档

  type => 'blanks'
    The "blanks" type is used to highlight blank cells in a range:
	"blanks"����������һ����Χ�ڸ����հ׵�Ԫ����

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'blanks',
                format   => $format,
            }
        );

  type => 'no_blanks'
    The "no_blanks" type is used to highlight non blank cells in a range:
	"no_blanks"����������һ����Χ�ڸ����ǿհ׵�Ԫ����

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'no_blanks',
                format   => $format,
            }
        );

  type => 'errors'
    The "errors" type is used to highlight error cells in a range:
	"errors"����������һ����Χ�ڸ����д����ĵ�Ԫ����

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'errors',
                format   => $format,
            }
        );

  type => 'no_errors'
    The "no_errors" type is used to highlight non error cells in a range:
	"no_errors"����������һ����Χ�ڸ���û�д����ĵ�Ԫ����

        $worksheet->conditional_formatting( 'A1:A4',
            {
                type     => 'no_errors',
                format   => $format,
            }
        );

  type => '2_color_scale'
    The "2_color_scale" type is used to specify Excel's "2 Color Scale"
    style conditional format.
	"2_color_scale"��������ָ��Excel��"2 Color Scale"������������ʽ��

        $worksheet->conditional_formatting( 'A1:A12',
            {
                type  => '2_color_scale',
            }
        );

    At the moment only the default colors and properties can be used. These
    will be extended in time.
	����ֻ��ʹ��Ĭ�ϵ���ɫ�����ԡ������ἰʱ������չ��

  type => '3_color_scale'
    The "3_color_scale" type is used to specify Excel's "3 Color Scale"
    style conditional format.
	
    "3_color_scale"��������ָ��Excel��"3 Color Scale"������������ʽ��
        $worksheet->conditional_formatting( 'A1:A12',
            {
                type  => '3_color_scale',
            }
        );

    At the moment only the default colors and properties can be used. These
    will be extended in time.
	
	����ֻ��ʹ��Ĭ�ϵ���ɫ�����ԡ������ἰʱ������չ��

  type => 'data_bar'
    The "data_bar" type is used to specify Excel's "Data Bar" style
    conditional format.
    "data_bar"��������ָ��Excel��"data_bar"������������ʽ��
        $worksheet->conditional_formatting( 'A1:A12',
            {
                type  => 'data_bar',
            }
        );

    At the moment only the default colors and properties can be used. These
    will be extended in time.
    ����ֻ��ʹ��Ĭ�ϵ���ɫ�����ԡ������ἰʱ������չ��
  type => 'formula'
    The "formula" type is used to specify a conditional format based on a
    user defined formula:
	"formula"�������ڸ����û������Ĺ�ʽָ��һ��������ʽ��

    $worksheet->conditional_formatting( 'A1:A4', { type => 'formula',
    criteria => '=$A$1 > 5', format => $format, } );

    The formula is specified in the "criteria".
	��ʽ��"criteria"��ָ����

  Conditional Formatting Examples
  ������ʽ����
    Example 1. Highlight cells greater than or equal to an integer value.
	��1.����ֵ���ڻ�����ĳ������ֵ�ĵ�Ԫ����

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
	��2.����ֵ���ڻ�����ĳ��ֵ�����õ�Ԫ����


        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'cell',
                criteria => 'greater than',
                value    => '$H$1',
                format   => $format,
            }
        );

    Example 3. Highlight cells greater than a certain date:
	��3.������ֵ��ĳһȷ�����ڴ��ĵ�Ԫ����

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'date',
                criteria => 'greater than',
                value    => '2011-01-01T',
                format   => $format,
            }
        );

    Example 4. Highlight cells with a date in the last seven days:
	��4.������������7�����ڵĵ�Ԫ����

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'time_period',
                criteria => 'last 7 days',
                format   => $format,
            }
        );

    Example 5. Highlight cells with strings starting with the letter "b":
	��5.�����ַ��������ַ�"b"��ͷ�ĵ�Ԫ����

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
	��6.����һ����Χ�ڣ�������׼����ƽ������1�ĵ�Ԫ����

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'average',
                format   => $format,
            }
        );

    Example 7. Highlight duplicate cells in a range:
	��7.����һ����Χ����ȫ��ͬ�ĵ�Ԫ����

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'duplicate',
                format   => $format,
            }
        );

    Example 8. Highlight unique cells in a range.
	��8.������һ����Χ��Ψһ�ĵ�Ԫ����

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'unique',
                format   => $format,
            }
        );

    Example 9. Highlight the top 10 cells.
	��9.����ͷ10����Ԫ����

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'top',
                value    => 10,
                format   => $format,
            }
        );

    Example 10. Highlight blank cells.
	��10.�����հ׵�Ԫ����

        $worksheet->conditional_formatting( 'A1:F10',
            {
                type     => 'blanks',
                format   => $format,
            }
        );

    Ҳ���鿴 "conditional_format.pl"������

   Excel�еĹ�ʽ�ͺ���
  Introduction ����
  
	������Excel��Excel::Writer::XLSX�й�ʽ�ͺ����ļ������ܡ�

	��ʽ����һ���Ⱥſ�ʼ���ַ�����

        '=A1+B1'
        '=AVERAGE(1, 2, 3)'

    The formula can contain numbers, strings, boolean values, cell
    references, cell ranges and functions. Named ranges are not supported.
    Formulas should be written as they appear in Excel, that is cells and
    functions must be in uppercase.
	��ʽ���԰������֡��ַ���������ֵ����Ԫ�����á���Ԫ��ֵ���ͺ�������֧������ֵ�򡣹�ʽӦ������Excel������д�룬����Ԫ���ͺ��������Ǵ�д�ġ�

    Cells in Excel are referenced using the A1 notation system where the
    column is designated by a letter and the row by a number. Columns range
    from A to XFD i.e. 0 to 16384, rows range from 1 to 1048576. 
	��Ԫ��ʹ��A1��ʾ��ϵͳ���ã�������ĸ��ʾ�У����ֱ�ʾ�С��еķ�Χ�Ǵ�A��XFD��0..16384���еķ�Χ�Ǵ�1��1048576.���磺

        use Excel::Writer::XLSX::Utility;

        ( $row, $col ) = xl_cell_to_rowcol( 'C2' );    # (1, 2)
        $str = xl_rowcol_to_cell( 1, 2 );              # C2

    The Excel "$" notation in cell references is also supported. This allows
    you to specify whether a row or column is relative or absolute. This
    only has an effect if the cell is copied. The following examples show
    relative and absolute values.
	��Excel�Ҳ֧�ֵ�Ԫ�����õ�"$"��ʾ������������ָ���л����Ƿ����������û�����Ӧ�á���ֻ���ڸ��Ƶ�Ԫ��ʱ�����á�������������ʾ�����Ժ;���ֵ��

        '=A1'   # �к��������Ե�
        '=$A1'  # ��ʽ���Եģ��������Ե�
        '=A$1'  # �������Եģ����Ǿ��Ե�
        '=$A$1' # �к����Ǿ��Ե�

	��ʽҲ�����õ�ǰ�������������������еĵ�Ԫ�������磺

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
	���������ú͵�Ԫ�����ñ���̾�ŷ��롣���������������пո񡢶��Ż����ţ���ExcelҪ��ʹ�õ����Ž���������������������������2������һ����  
	
    Ϊ����ʹ��̫��ת���ַ���������ʹ������������"q{}"�������š�ֻ��ʹ��"add_worksheet()"�������ӵĺϷ������������ܱ����ڹ�ʽ���㲻�������ⲿ��������
	
    The following table lists the operators that are available in Excel's
    formulas. The majority of the operators are the same as Perl's,
    differences are indicated:
	
	�����ı��г�����Excel��ʽ�п��õĲ����������д�������������Perl�еĲ�������ͬ����֮ͬ��Ҳ��ָ�����ˣ�

        ����������:
        =====================
        ������    ����                      ����
           +      Addition                  1+2
           -      Subtraction               2-1
           *      Multiplication            2*3
           /      Division                  1/4
           ^      Exponentiation            2^3      # �ȼ��� **
           -      Unary minus               -(1+2)   # ����֧��
           %      Percent (Not modulus)     13%      # ��֧��, [1]


        �Ƚϲ�����:
        =====================
        Operator  Meaning                   Example
            =     Equal to                  A1 =  B1 #�ȼ���==
            <>    Not equal to              A1 <> B1 #�ȼ���!=
            >     Greater than              A1 >  B1
            <     Less than                 A1 <  B1
            >=    Greater than or equal to  A1 >= B1
            <=    Less than or equal to     A1 <= B1


        �ַ���������:
        ================
        Operator  Meaning                   Example
            &     Concatenation             "Hello " & "World!" # [2]


        Reference operators:
        ====================
        Operator  Meaning                   Example
            :     Range operator            A1:A4               # [3]
            ,     Union operator            SUM(1, 2+2, B3)     # [4]


        ע��:
		[1]:You can get a percentage with formatting and modulus with MOD().
        [1]: ������ʹ�ø�ʽ���õ�һ���ٷ�����ʹ��MOD()�õ�һ��ģ��
		[2]: ��Perl����("Hello " . "World!")�ȼۡ�
		[3]: �÷�Χ�ȼ��ڵ�Ԫ�� A1, A2, A3�� A4.
        [4]: ������Perl�е��б���������Ϊ���ơ�

    The range and comma operators can have different symbols in non-English
    versions of Excel. These will be supported in a later version of
    Excel::Writer::XLSX. European users of Excel take note:
	��Χ�Ͷ��Ų������ڷ�Ӣ���汾��Excel���в�ͬ�ķ��š���Щ�����Ժ��汾��Excel::Writer::XLSX��֧�֡�ŷ�޵�Excel�û�ע�⣺

        $worksheet->write('A1', '=SUM(1; 2; 3)'); # Wrong!!
        $worksheet->write('A1', '=SUM(1, 2, 3)'); # Okay

   
    If your formula doesn't work in Excel::Writer::XLSX try the following:
	�������Ĺ�ʽ��Excel::Writer::XLSX�в������ã����������ģ�

        1. Verify that the formula works in Excel (or Gnumeric or OpenOffice.org).
		1.���鹫ʽ��Excel����Ч
        2. Ensure that cell references and formula names are in uppercase.
		2.ȷ����Ԫ�����ú͹�ʽ�����Ǵ�д�ġ�
        3. Ensure that you are using ':' as the range operator, A1:A4.
		3.ȷ����ʹ��':'��Ϊ��Χ��������A1:A4.
        4. Ensure that you are using ',' as the union operator, SUM(1,2,3).
		4.ȷ����ʹ��','��Ϊ������������SUM(1,2,3).
        5. Ensure that the function is in the above table.
		5.ȷ���������ϱ��г��ֵĺ�����

   

EXAMPLES ����
    �鿴 Excel::Writer::XLSX::Examples ��ȡ������ʾ���嵥.

  ��1
    The following example shows some of the basic features of
    Excel::Writer::XLSX.
	������������ʾ�� Excel::Writer::XLSX��һЩ����������

        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        # ����һ����simple.xlsx���¹�����������һ�Ź�����
        my $workbook  = Excel::Writer::XLSX->new( 'simple.xlsx' );
        my $worksheet = $workbook->add_worksheet();

        # The general syntax is write($row, $column, $token). Note that row and
        # column are zero indexed

        # д��һЩ�ı�
        $worksheet->write( 0, 0, 'Hi Excel!' );


        # Write some numbers
        $worksheet->write( 2, 0, 1 );
        $worksheet->write( 3, 0, 1.00000 );
        $worksheet->write( 4, 0, 2.00001 );
        $worksheet->write( 5, 0, 3.14159 );


        # Write some formulas
        $worksheet->write( 7, 0, '=A3 + A6' );
        $worksheet->write( 8, 0, '=IF(A5>3,"Yes", "No")' );


        # д��һ��������
        $worksheet->write( 10, 0, 'http://www.perl.com/' );

  ��2
    The following is a general example which demonstrates some features of
    working with multiple worksheets.
	������һЩ��ͨ�����ӣ�����˵����һЩʹ�ö��Ź����������ԡ�

        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        # ����һ���µ�Excel������
        my $workbook = Excel::Writer::XLSX->new( 'regions.xlsx' );

        # ����һЩ������
        my $north = $workbook->add_worksheet( 'North' );
        my $south = $workbook->add_worksheet( 'South' );
        my $east  = $workbook->add_worksheet( 'East' );
        my $west  = $workbook->add_worksheet( 'West' );

        # ����һ����ʽ
        my $format = $workbook->add_format();
        $format->set_bold();
        $format->set_color( 'blue' );

        #��ÿ�Ź���������һ������
        for my $worksheet ( $workbook->sheets() ) {
            $worksheet->write( 0, 0, 'Sales', $format );
        }

        # д��һЩ����
        $north->write( 0, 1, 200000 );
        $south->write( 0, 1, 100000 );
        $east->write( 0, 1, 150000 );
        $west->write( 0, 1, 100000 );

        # ���û������
        $south->activate();

        # ���õ�һ�еĿ���
        $south->set_column( 0, 0, 20 );

        # ���û��Ԫ��
        $south->set_selection( 0, 1 );

  ��3
    Example of how to add conditional formatting to an Excel::Writer::XLSX
    file. The example below highlights cells that have a value greater than
    or equal to 50 in red and cells below that value in green.
	������������һ��Excel::Writer::XLSX��ʽ���ļ�����������ʽ�����ӡ�����������ʹ�ú�ɫ����ֵ���ڻ�����50�ĵ�Ԫ����ʹ����ɫ����ֵС��50�ĵ�Ԫ����

        #!/usr/bin/perl

        use strict;
        use warnings;
        use Excel::Writer::XLSX;

        my $workbook  = Excel::Writer::XLSX->new( 'conditional_format.xlsx' );
        my $worksheet = $workbook->add_worksheet();

		#����������ʹ�ú�ɫ����ֵ���ڻ�����50�ĵ�Ԫ����ʹ����ɫ����ֵС��50�ĵ�Ԫ��
		
        # Light red fill with dark red text.
		#ʹ�ð���ɫ�ı����е���ɫ���䣿
        my $format1 = $workbook->add_format(
            bg_color => '#FFC7CE',
            color    => '#9C0006',

        );

        # Green fill with dark green text.
		#ʹ�ð���ɫ�ı�������ɫ����
        my $format2 = $workbook->add_format(
            bg_color => '#C6EFCE',
            color    => '#006100',

        );

        # Some sample data to run the conditional formatting against.
		#һЩ�򵥵���������������ʽ������
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

        # д������.
        $worksheet->write( 'A1', $caption );
        $worksheet->write_col( 'B3', $data );

        # ��һ����Ԫ����Χ��д��������ʽ.
        $worksheet->conditional_formatting( 'B3:K12',
            {
                type     => 'cell',
                criteria => '>=',
                value    => 50,
                format   => $format1,
            }
        );

        # ��ͬһ��Ԫ����Χ��д������һ��������ʽ
        $worksheet->conditional_formatting( 'B3:K12',
            {
                type     => 'cell',
                criteria => '<',
                value    => 50,
                format   => $format2,
            }
        );

  Example 4
	����������ʹ�ú�����һ���򵥵����ӣ�

        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        # ����һ���µĹ�����������һ�Ź�����
        my $workbook  = Excel::Writer::XLSX->new( 'stats.xlsx' );
        my $worksheet = $workbook->add_worksheet( 'Test data' );

        # Ϊ��һ�������п�Ϊ
        $worksheet->set_column( 0, 0, 20 );


        # �����ⴴ��һ����ʽ
        my $format = $workbook->add_format();
        $format->set_bold();


        # д����������
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

        # д��һЩͳ�ƺ���
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
    �����������ӽ�һ��ʹ��tab�ַ��ָ��Ľ�����tab.txt�����ļ�ת��Ϊһ������"tab.xlsx"��Excel�ļ���

        #!/usr/bin/perl -w

        use strict;
        use Excel::Writer::XLSX;

        open( TABFILE, 'tab.txt' ) or die "tab.txt: $!";

        my $workbook  = Excel::Writer::XLSX->new( 'tab.xlsx' );
        my $worksheet = $workbook->add_worksheet();

        # �к��д�0��ʼ����
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

 	ע�⣺��ֻ��һ��˵���Եļ򵥵�ת��������ת��CSV�ļ���ʹ��Tab���ŷָ��������κ���ʽ�ķָ����ָ����ı��ļ������Ƽ������ܵ�csv2xls����������Text::CSV_XSģ����һ���֡�

   �ڴ˴��鿴 examples/csv2xls ����:
    <http://search.cpan.org/~hmbrand/Text-CSV_XS/MANIFEST>.

  ���ӵ�����
    ������Excel::Writer::XLSX��׼���а����ṩ��ʵ�������ļ�������˵���˸�ģ���Ĳ�ͬ������ѡ��鿴Excel::Writer::XLSX::Examples��ȡ����ϸ����Ϣ.
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

LIMITATIONS ����
    The following limits are imposed by Excel 2007+:

        ����                                 ����        -----------------------------------  ------
        һ���ַ����е������ַ���             32,767
        ����������                           16,384
        ����������                           1,048,576
        ���������е������ַ���               31
        ҳü/ҳ���е������ַ���              254

��Spreadsheet::WriteExcelģ���ļ�����
	"Excel::Writer::XLSX"ģ���� "Spreadsheet::WriteExcel"ģ����������

	��֧�� Spreadsheet::WriteExcel���е����ԣ�ע������΢С�Ĳ�ͬ��

        ����������                  ֧��
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
        set_optimization()          Yes. Spreadsheet::WriteExcel�в���Ҫ.
        add_chart_ext()             Not supported.Excel::Writer::XLSX�в��Ǳ�����
        compatibility_mode()        Deprecated. Excel::Writer::XLSX�в��Ǳ�����
        set_codepage()              Deprecated. Excel::Writer::XLSX�в��Ǳ�����


        Worksheet Methods           Support
        =================           =======
        write()                     Yes
        write_number()              Yes
        write_string()              Yes
        write_rich_string()         Yes. Spreadsheet::WriteExcel��û�и÷���.
        write_blank()               Yes
        write_row()                 Yes
        write_col()                 Yes
        write_date_time()           Yes
        write_url()                 Yes
        write_formula()             Yes
        write_array_formula()       Yes.Spreadsheet::WriteExcel��û�и÷���.
        keep_leading_zeros()        Yes
        write_comment()             Yes
        show_comments()             Yes
        set_comments_author()       Yes
        add_write_handler()         Yes
        insert_image()              Yes/Partial, �鿴�ĵ�.
        insert_chart()              Yes
        data_validation()           Yes
        conditional_format()        Yes. Spreadsheet::WriteExcel��û�и÷���.
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
        merge_range_type()          Yes. Spreadsheet::WriteExcel��û�и÷�����
        set_zoom()                  Yes
        right_to_left()             Yes
        hide_zero()                 Yes
        set_tab_color()             Yes
        autofilter()                Yes
        filter_column()             Yes
        filter_column_list()        Yes. Spreadsheet::WriteExcel��û�и÷���.
        write_utf16be_string()      ���Ƽ�ʹ��. ʹ�� Perl utf8�ַ�������.
        write_utf16le_string()      ���Ƽ�ʹ��. ʹ�� Perl utf8�ַ�������.
        store_formula()             ���Ƽ�ʹ��. �鿴�ĵ�.
        repeat_formula()            ���Ƽ�ʹ��. �鿴�ĵ�.
        write_url_range()           Not supported. Excel::Writer::XLSX�в��Ǳ�����

        ҳ�����÷���                ֧��
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

        ��ʽ����                    ֧��
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

REQUIREMENTS  Ҫ��
    <http://search.cpan.org/search?dist=Archive-Zip/>.

    Perl 5.10.0.

SPEED AND MEMORY USAGE  �ٶȺ��ڴ�ʹ��

	"Spreadsheet::WriteExcel"�����Ż��ٶȲ������ڴ�ʹ�á�����������Ŀ����ζ��ʵ�������û�Ҫ����������ʽ���͵�����д�����ݹ��ܲ������ס�

     ���ˣ�"Excel::Writer::XLSX"��ȡ��ͬ�����Ʒ����������ڴ��д������������ݣ������ڹ��ܸ����ӡ������Ľ�������Excel::Writer::XLSX ��Spreadsheet::WriteExcel ����50%������������ʹ�ø����ڴ档�������Ӹ������к��з�Χʱ��������Ϊ���������ļ����ù��ڴ档����Spreadsheet::WriteExcel���������������⡣
   
	ʹ�ù�������"set_optimization()"�����������ڴ�ʹ�ü�������ȫ��С��

        $workbook->set_optimization();

 
	�������Ĵ��۾����㲻��ʹ���κβ�����Ԫ�����¹��ܣ�д�����ݺ󣬸��Ż�ѡ��򿪡�


DIAGNOSTICS  ����
    Filename required by Excel::Writer::XLSX->new()
       ���������캯��һ���ļ�����

    Can't open filename. It may be in use or protected.
       
		���ܴ����ļ�д�롣��Ҫд�����Ǹ��ļ��п��ܱ�д�������ļ�������������ʹ�á�

    Can't call method "XXX" on an undefined value at someprogram.pl.
      
    ��Windows���ⳣ���������������Դ������ļ���һ���Ѿ���Excel�򿪲���ס���ļ���ͻ�ˡ�
    The file you are trying to open 'file.xls' is in a different format than
    specified by the file extension.

		���㴴��һ��XLSX�ļ������Ǹ���һ��xls��׺ʱ�������ָþ��档

WRITING EXCEL FILES
    
    �����������󣬱�����һ���ĸо��������ܸ�ϲ�������ķ���֮һ������д��Excel��
    *   Spreadsheet::WriteExcel

 
		����Excel::Writer::XLSX ����������ʹ��ͬ���Ľӿڡ�������xls��ʽ���ļ�����Excel 97-2003�汾����Щ�ļ���Ȼ�ܱ�Excel2007��ȡ������֧�ֵ��к��������������ơ�

    *   Win32::OLE module and office automation

         ����Ҫһ��Windowsƽ̨����װһ��Excel������������Excel��������ǿ������ȫ�ķ�����
    *   CSV, comma separated variables or text

        �����ļ���չ����csv��Excel�򿪺����Զ�ת���ø�ʽ������һ��CSV�ļ�����������������ô���ס��鿴DBD::RAM, DBD::CSV, Text::xSV �� Text::CSV_XS ģ�顣
    *   DBI with DBD::ADO or DBD::ODBC

        Excel�ļ������ڲ����������������Ǳ��ֵ���һ�����ݿ⡣ʹ�ñ�׼��Perl���ݿ�ģ��֮һ�������Խ�һ��Excel�ļ��������ݿ����ӡ�

    *   DBD::Excel

         
		��Ҳ����ͨ�� DBD::Excel ģ��ʹ�ñ�׼��DBI�ӿڷ���Spreadsheet::WriteExcel ��
		�鿴  <http://search.cpan.org/dist/DBD-Excel>
    *   Spreadsheet::WriteExcelXML
       
        ��ģ��������ʹ���� Spreadsheet::WriteExcelͬ���Ľӿ�������Excel XML �ļ���
		�鿴 <http://search.cpan.org/dist/Spreadsheet-WriteExcelXML>.
		
    *   Excel::Template

        ��ģ������������ĳ���������� HTML::Template���Ƶ�XML ģ���ϴ����ļ���
		�鿴<http://search.cpan.org/dist/Excel-Template/>.
        

    *   Spreadsheet::WriteExcel::FromXML

       
		��ģ��������ʹ��Spreadsheet::WriteExcel ��Ϊ��̨��һ���򵥵�XML�ļ�ת��ΪExcel�ļ���
		XML�ĸ�ʽ��֧�ֵ�DTD�����塣
		�鿴<http://search.cpan.org/dist/Spreadsheet-WriteExcel-FromXML>.

    *   Spreadsheet::WriteExcel::Simple

       ���ṩ�˶�Spreadsheet::WriteExcel���Ӽ򵥵Ľӿڡ�
        <http://search.cpan.org/dist/Spreadsheet-WriteExcel-Simple>.

    *   Spreadsheet::WriteExcel::FromDB

  	 �����ڴ�DB���д���Excel�ļ������á�
	 <http://search.cpan.org/dist/Spreadsheet-WriteExcel-FromDB>.

    *   HTML tables

        This is an easy way of adding formatting via a text based format.
		ͨ��һ�������ı��ĸ�ʽ���Ӹ�ʽ�����ס�

    *   
READING EXCEL FILES  ��ȡExcel�ļ�
	��Excel�ж�ȡ���ݣ����ԣ�

    *   Spreadsheet::ParseExcel

        
        ��ʹ��OLE::Storage-Liteģ����Excel����ȡ���ݡ�
		�鿴<http://search.cpan.org/dist/Spreadsheet-ParseExcel>.
    *   Spreadsheet::ParseExcel_XLHTML


    *   XML::Excel
         ʹ��Spreadsheet::ParseExcel��Excel�ļ�ת��ΪXML�ļ�
        �鿴��<http://search.cpan.org/dist/XML-Excel>.

    *   
