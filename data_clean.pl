use Mojo::DOM;
use 5.010;
use Data::Dumper;
use Encode;
use Excel::Writer::XLSX;


my $workbook  = Excel::Writer::XLSX->new("data_clean.xlsx");
my $worksheet = $workbook->add_worksheet('data_clean');
my $content = do {local $/;<DATA>};
$content =~ s/&nbsp;//g;
my $dom = Mojo::DOM->new($content);
my $headers = $dom->find('table[id="tabGoodsInfo"]');
my @columns;
my @headers; 

# 收集产品信息到数组中
for my $e ($headers->find('table[id="tabGoodsInfo"]> tr > td > font > input[class="inputnormal"]')->each) {
  # say $e; 
  # 21000101 通气鼻贴 10片 透明型（标准） 器械 14.4900 19.5000
  # 箱 中盒 盒 天津 中美天津史克制药有限公司  2014-05-11～2014-06-19 中美史克
	my $collect= Mojo::DOM->new($e);
	$collect->find('input[value]')->each(sub { push @headers,shift->{value}});
}

# 产品的每条记录
my $collection = $dom->find('table[id="dgGoods"]>tr');
 foreach my $ele (@$collection) { 
     my $tds = $ele->find('td')->text;
	 next if $tds =~ /门店代码/;
	 next if $tds =~ /合计/;
	 push @columns,[@$tds,@headers];
}

foreach my $line (@columns){
          
    foreach (@$line) {
	    chomp;
        $_=decode("gb2312",$_) if defined $_;
    }
}

my @first_line = (
                     ['门店代码',	'门店名称',	'销售日期',	'批号',	'数量',	'供应价',
					 '代码',       '品名',        '规格',        '剂型',    '批发价',  '零售价',
					 '外包装',     '中包装',      '内包装',      '产地',    '厂家',    '查询时间', '业务单位']
					 );
					 
foreach my $line (@first_line){
          
    foreach (@$line) {
	    chomp;
        $_=decode("gb2312",$_) if defined $_;
    }
}

$worksheet->write_col('A1',\@first_line);	# 写入标题	 
$worksheet->write_col('A2',\@columns);


__DATA__
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<HTML>

	<HEAD>

		<title>SaleRecords</title>

		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">

		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">

		<meta name="vs_defaultClientScript" content="JavaScript">

		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">

		<LINK href="../BC.css" type="text/css" rel="stylesheet">

		<link rel="stylesheet" href="../menu.css">

	</HEAD>

	<body bottomMargin="0" leftMargin="0" topMargin="0">

		<form name="Form1" method="post" action="SaleRecords.aspx" id="Form1">

<input type="hidden" name="__VIEWSTATE" id="__VIEWSTATE"  />



<input type="hidden" name="__EVENTVALIDATION" id="__EVENTVALIDATION"  />

			<TABLE class="table" id="Table0" style="LEFT: 4px; WIDTH: 100%; TOP: 0px; HEIGHT: 100%"

				cellSpacing="0" cellPadding="0" align="center" border="0">

				<TR>

					<TD>

<link href="BC.css" type="text/css" rel="stylesheet">

<link rel="stylesheet" href="menu.css">



<script language="JavaScript">

<!-- Hide

var timerID = null

var timerRunning = false

function MakeArray(size) {

    this.length = size;

    for(var i = 1; i <= size; i++)

      {

        this[i] = "";

      }

  return this;

}

function stopclock (){

    if(timerRunning)

    clearTimeout(timerID);

    timerRunning = false

}

function showtime (){

  var now = new Date();

  var year = now.getYear();

  var month = now.getMonth() + 1;

  var date = now.getDate();

  var hours = now.getHours();

  var minutes = now.getMinutes();

  var seconds = now.getSeconds();

  var day = now.getDay();

  Day = new MakeArray(7);

  Day[0]="星期天";

  Day[1]="星期一";

  Day[2]="星期二";

  Day[3]="星期三";

  Day[4]="星期四";

  Day[5]="星期五";

  Day[6]="星期六";

  var timeValue = "";

  timeValue += year + "年";

  timeValue += ((month < 10) ? "0" : "") + month + "月";

  timeValue += date + "日  ";

  timeValue += (Day[day]) + "  ";

  timeValue += ((hours <= 12) ? hours : hours - 12);

  timeValue += ((minutes < 10) ? ":0" : ":") + minutes;

  timeValue += ((seconds < 10) ? ":0" : ":") + seconds;

  timeValue += (hours < 12) ? "上午" : "  下午";

  document.all.clock.innerText = timeValue;

  timerID = setTimeout("showtime()",1000);

  timerRunning = true

}

function startclock () {

  stopclock();

  showtime()

}

//-->

</script>

<table border="0" cellpadding="0" cellspacing="0" style="Z-INDEX: 101; LEFT: 0px; TOP: 0px; BORDER-COLLAPSE: collapse"

	bordercolor="#111111" width="100%" bgcolor="#336699">

	<tr>

		<td>

			<table border="0" cellpadding="0" cellspacing="0" style="Z-INDEX: 101; LEFT: 16px; BORDER-COLLAPSE: collapse"

				bordercolor="#111111" width="100%" bgcolor="#336699">

				<tr>

					<td width="500" colSpan="1" rowSpan="1" height="58">

						<TABLE id="Table1" cellSpacing="0" cellPadding="0" border="0" width="100%" style="HEIGHT: 58px">

							<TR>

								<TD background="/BC/image/yht.jpg" style="WIDTH: 70px; HEIGHT: 15px"><FONT face="宋体"></FONT></TD>

								<td style="FONT-WEIGHT: bolder; FONT-SIZE: x-large; WIDTH: 500px; COLOR: white; FONT-FAMILY: 隶书; HEIGHT: 15px"><FONT face="宋体">业务单位信息查询系统</FONT></td>

							</TR>

						</TABLE>

					</td>

					<td>

						<!--

						<table cellSpacing="0" cellPadding="0" border="0" width="100%" style="HEIGHT: 58px">

							<tr>

								<td valign="bottom" width="33%" height="58">

									<p align="center"><a class="toolbar" href="/BC/Page/BrowseCompany.aspx">单位信息</a></p>

								</td>

								<td valign="bottom" width="2" height="58">

									<p align="center">

										<img height="14" width="2" border="0" src="/BC/Image/separator.gif"></p>

								</td>

								<td valign="bottom" width="33%" height="58">

									<p align="center"><a class="toolbar" href="/BC/Page/SaleRecords.aspx">销售查询</a></p>

								</td>

								<td valign="bottom" width="2" height="58">

									<p align="center">

										<img height="14" width="2" border="0" src="/BC/Image/separator.gif"></p>

								</td>

								<td valign="bottom" width="34%" height="58">

									<p align="center"><a class="toolbar" href="/BC/Help/Instruction_BC.mht"  target=_blank>使用指南</a></p>

								</td>

							</tr>

						</table>

						-->

					</td>

				</tr>

			</table>

		</td>

	</tr>

	<tr>

		<td>

			<table border="0" cellpadding="0" cellspacing="0" style="WIDTH: 100%; COLOR: white; BORDER-COLLAPSE: collapse; HEIGHT: 32px"

				bordercolor="#111111" height="32" bgcolor="#6699cc">

				<tr>

					<td align="left" valign="bottom" width="34%">欢迎您：<a id="Title1_UserInfo" class="toolbar">中美史克</a>

					</td>

					<td align="center" valign="bottom" width="11%"><a class="toolbar" href="/BCC/page/Login.aspx">重新登录</a></td>

					<td align="center" valign="bottom" width="11%"><a class="toolbar" href="/BCC/page/ChangePWD.aspx">修改密码</a>

					</td>

					<td width="14%">&nbsp;</td>

					<td id="clock" align="right" valign="bottom" width="30%"></td>

				</tr>

			</table>

		</td>

	</tr>

</table>

<script language="javascript">

	startclock();

</script>

<!-- menu script itself. you should not modify this file -->

<script type="text/javascript" language="javascript" src="/BCC/JS/menucode.js"></script>

<script type="text/javascript" language="JavaScript" src="/BCC/JS/menu.js"></script>

<!-- items structure. menu hierarchy and links are stored there -->

<script language="JavaScript" src="/BCC/JS/menu_items.js"></script>

<!-- files with geometry and styles structures -->

<script language="JavaScript" src="/BCC/JS/menu_tpl.js"></script>

<script language="JavaScript">

	<!--//

	// Note where menu initialization block is located in HTML document.

	// Don't try to position menu locating menu initialization block in

	// some table cell or other HTML element. Always put it before </body>



	// each menu gets two parameters (see demo files)

	// 1. items structure

	// 2. geometry structure



	new menu (MENU_ITEMS, MENU_POS);

	// make sure files containing definitions for these variables are linked to the document

	// if you got some javascript error like "MENU_POS is not defined", then you've made syntax

	// error in menu_tpl.js file or that file isn't linked properly.

	

	// also take a look at stylesheets loaded in header in order to set styles

	//-->

</script></TD>

				</TR>

				<TR>

					<TD vAlign="top" align="center" height="100%">

						<TABLE class="table" id="Table1" cellSpacing="0" cellPadding="0" width="100%" align="center"

							border="0">

							<TR>

								<TD colSpan="2" height="5"></TD>

							</TR>

							<TR>

								<TD class="titlecell" style="HEIGHT: 25px" align="left"><B>查询条件</B></TD>

								<TD class="titlecell" style="HEIGHT: 25px" align="right"><input type="submit" name="btFind" value="查询" onclick="javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions(&quot;btFind&quot;, &quot;&quot;, true, &quot;&quot;, &quot;&quot;, false, false))" id="btFind" class="button" /></TD>

							</TR>

							<TR>

								<td class="errorcellfore" colSpan="12"><span id="lblError"></span>

									

									

									</td>

							</TR>

							<TR>

								<td colSpan="2" style="HEIGHT: 1px">

									<TABLE class="table" id="Table3" cellSpacing="0" cellPadding="0" width="100%" align="center"

										border="1">

										<tr>

											<TD style="HEIGHT: 23px" align="center" width="70">开始时间</TD>

											<TD style="HEIGHT: 23px" width="80"><input name="txtBeginDate" type="text" value="2014-05-11" id="txtBeginDate" title="格式例：2005年1月1日，可输入：05-1-1或05.1.1" class="inputnormal" /></TD>

											<TD style="HEIGHT: 23px" align="center" width="70">结束时间</TD>

											<TD style="HEIGHT: 23px" width="80"><input name="txtEndDate" type="text" value="2014-06-19" id="txtEndDate" class="inputnormal" /></TD>

											<TD style="HEIGHT: 23px" align="center" width="70">业务单位</TD>

											<TD style="HEIGHT: 23px" width="320"><FONT face="宋体"><select name="ddlCompany" id="ddlCompany" class="inputnormal">

	<option selected="selected" value="002-02-001">[002-02-001]中美史克</option>



</select></FONT></TD>

											<TD style="WIDTH: 50px; HEIGHT: 23px" align="center">品种</TD>

											<TD style="WIDTH: 320px; HEIGHT: 23px"><FONT face="宋体"><select name="ddlGoods" id="ddlGoods" class="inputnormal">

	<option value="">--选择--</option>

	<option value="12300003">[12300003](F)丙酸氟替卡松鼻喷雾剂(辅舒良)50ug/喷&#215;120喷/瓶,琥珀色玻璃瓶</option>

	<option value="14440140">[14440140]对乙酰氨基酚片(必理通)500mg&#215;10片</option>

	<option value="14440282">[14440282]布洛芬缓释胶囊(芬必得)300mg&#215;20粒/盒,铝塑泡罩</option>

	<option value="14440286">[14440286]布洛芬缓释胶囊(芬必得)0.3g*10s</option>

	<option value="14440289">[14440289]布洛芬缓释胶囊0.4g*8粒/板*3板/盒</option>

	<option value="14440441">[14440441]酚咖片(芬必得)10片/盒,PTP铝塑/聚氯乙烯/聚偏二氯乙烯硬片泡罩</option>

	<option value="14440442">[14440442]酚咖片(芬必得)565mg&#215;20片/盒,铝塑泡罩</option>

	<option value="14440493">[14440493](FM)布洛伪麻缓释胶囊(康泰克清)10s</option>

	<option value="14440580">[14440580](FM)复方盐酸伪麻黄碱缓释胶囊(新康泰克)10粒&#215;1盒/盒,铝塑泡罩</option>

	<option value="14440582">[14440582](FM)复方盐酸伪麻黄碱缓释胶囊8粒/盒</option>

	<option value="14440589">[14440589](FM)美扑伪麻片(新康泰克)10片&#215;2板/盒,铝塑</option>

	<option value="14440590">[14440590](FM)美扑伪麻片10片/盒,铝塑泡罩</option>

	<option value="14450025">[14450025]西咪替丁片(泰胃美)800mg&#215;10片/盒,铝塑泡罩,(薄膜衣片)</option>

	<option value="14450026">[14450026]西咪替丁片(泰胃美)400mg&#215;20片/盒,铝塑泡罩,(薄膜衣片)</option>

	<option value="14468020">[14468020]阿苯达唑片(史克肠虫清)200mg&#215;10片/盒,铝塑泡罩</option>

	<option value="14468025">[14468025]阿苯达唑片2s*200mg</option>

	<option value="14510291">[14510291]盐酸氨溴索缓释胶囊75毫克*6粒/盒</option>

	<option value="15529011">[15529011](F)丙酸倍氯米松鼻气雾剂(伯克纳)50ug/喷&#215;200喷/支,铝罐包装</option>

	<option value="16541240">[16541240]莫匹罗星软膏(百多邦)2％/5g/支,复合铝管</option>

	<option value="16541241">[16541241]莫匹罗星软膏(百多邦)2％/10g/支,复合铝管</option>

	<option value="16541280">[16541280]布洛芬乳膏(芬必得)20g（5％）&#215;1支/支,铝管</option>

	<option value="16541330">[16541330]盐酸特比萘芬乳膏(兰美抒)5g&#215;1支/支,铝管</option>

	<option value="21000084">[21000084]通气鼻贴10片 肤色型（标准）</option>

	<option value="21000101">[21000101]通气鼻贴10片 透明型（标准）</option>

	<option value="21000136">[21000136]通气鼻贴(儿童型)8s</option>

	<option value="21000204">[21000204]通气鼻贴10片　薄荷型（标准）</option>

	<option value="42990002">[42990002]新康泰克喉爽润喉软糖（柠檬味）40克(20粒装)</option>

	<option value="42990003">[42990003]新康泰克喉爽草本润喉软糖（柠檬味）20克（10粒装）</option>

	<option value="42990004">[42990004]新康泰克喉爽草本润喉软糖（薄荷味）40克(20粒装)</option>

	<option value="42990005">[42990005]新康泰克喉爽草本润喉软糖（薄荷味）20克（10粒装）</option>

	<option value="42990119">[42990119]新康泰克喉爽草本润喉软糖（缤纷莓果味）40克（约20粒铁盒装）</option>

	<option value="43990000">[43990000]新康泰克草本润喉软糖喉爽（薄荷口味）40克（20粒装）</option>

	<option value="85001032">[85001032]百多邦创面消毒喷雾剂70ml</option>

	<option value="88000476">[88000476]舒适达抗敏感牙膏（全面护理）120g</option>

	<option selected="selected" value="88000477">[88000477]舒适达抗敏感牙膏（清新薄荷）120g</option>

	<option value="88000478">[88000478]舒适达速效抗敏牙膏120g</option>

	<option value="88000479">[88000479]舒适达专业修复牙膏100克</option>

	<option value="88009047">[88009047]保丽净假牙清洁片24片</option>

	<option value="88009049">[88009049]保丽净假牙清洁片（局部假牙专用）24片</option>

	<option value="88010004">[88010004]保丽净假牙清洁片（全/半口假牙专用）30片</option>

	<option value="88010005">[88010005]保丽净假牙清洁片（局部假牙专用）30片</option>



</select></FONT></TD>

										</tr>

									</TABLE>

								</td>

							</TR>

							<TR>

								<TD align="center" colSpan="2" height="5"></TD>

							</TR>

							<TR>

								<TD colspan="2" vAlign="top" align="center" height="100%">

									<table id="tabResult" class="table" cellspacing="0" cellpadding="0" width="100%" align="center" border="0">

	<tr>

		<TD class="titlecell" style="HEIGHT: 25px" align="left"><B>查找结果</B><B> </B>

												<span id="lblInfo" style="display:inline-block;">共找到 7 条记录</span></TD>

		<TD class="titlecell" style="HEIGHT: 25px" align="right"><input type="submit" name="btPtint" value="打印" onclick="javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions(&quot;btPtint&quot;, &quot;&quot;, true, &quot;&quot;, &quot;&quot;, false, false))" id="btPtint" class="button" />

												<input type="submit" name="btDown" value="下载" onclick="javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions(&quot;btDown&quot;, &quot;&quot;, true, &quot;&quot;, &quot;&quot;, false, false))" id="btDown" class="button" />

											</TD>

	</tr>

	<tr>

		<td colspan="2">

												<table id="tabGoodsInfo" class="table" style="TABLE-LAYOUT: fixed" height="100%" cellspacing="0" cellpadding="0" width="100%" align="center" border="1">

			<tr>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">代码</FONT></td>

				<td style="HEIGHT: 23px"><FONT face="宋体"><input name="txtCode" type="text" value="88000477" id="txtCode" class="inputnormal" /></FONT></td>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">品名</FONT></td>

				<td style="HEIGHT: 23px" colspan="3"><FONT face="宋体"><input name="txtName" type="text" value="舒适达抗敏感牙膏（清新薄荷）" id="txtName" class="inputnormal" /></FONT></td>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">规格</FONT></td>

				<td style="HEIGHT: 23px"><FONT face="宋体"><input name="txtSpec" type="text" value="120g" id="txtSpec" class="inputnormal" /></FONT></td>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">剂型</FONT></td>

				<td style="HEIGHT: 23px"><FONT face="宋体"><input name="txtDoseType" type="text" value="其他剂型" id="txtDoseType" class="inputnormal" /></FONT></td>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">批发价</FONT></td>

				<td style="HEIGHT: 23px"><FONT face="宋体"><input name="txtWholePrice" type="text" value="20.8100" id="txtWholePrice" class="inputnormal" /></FONT></td>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">零售价</FONT></td>

				<td style="HEIGHT: 23px"><FONT face="宋体"><input name="txtRetailPrice" type="text" value="28.0000" id="txtRetailPrice" class="inputnormal" /></FONT></td>

			</tr>

			<tr>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">外包装</FONT></td>

				<td style="HEIGHT: 23px"><FONT face="宋体"><input name="txtOutPack" type="text" value="箱" id="txtOutPack" class="inputnormal" /></FONT></td>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">中包装</FONT></td>

				<td style="HEIGHT: 23px"><FONT face="宋体"><input name="txtMidPack" type="text" value="封" id="txtMidPack" class="inputnormal" /></FONT></td>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">内包装</FONT></td>

				<td style="HEIGHT: 23px"><FONT face="宋体"><input name="txtInPack" type="text" value="支" id="txtInPack" class="inputnormal" /></FONT></td>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">产地</FONT></td>

				<td style="HEIGHT: 23px"><FONT face="宋体"><input name="txtProducePlace" type="text" value="江苏" id="txtProducePlace" class="inputnormal" /></FONT></td>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">厂家</FONT></td>

				<td style="HEIGHT: 23px" colspan="5"><FONT face="宋体"><input name="txtManufacturer" type="text" value="苏州克劳丽化妆品有限公司" id="txtManufacturer" class="inputnormal" /></FONT></td>

			</tr>

			<tr>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">查询时间</FONT></td>

				<td style="HEIGHT: 23px" colspan="7"><FONT face="宋体"><input name="txtDate" type="text" value="2014-05-11～2014-06-19" id="txtDate" class="inputnormal" /></FONT></td>

				<td style="HEIGHT: 23px" valign="middle" align="center" width="60"><FONT face="宋体">业务单位</FONT></td>

				<td style="HEIGHT: 23px" colspan="5"><FONT face="宋体"><input name="txtCompany" type="text" value="中美史克" id="txtCompany" class="inputnormal" /></FONT></td>

			</tr>

		</table>

		

											</td>

	</tr>

	<tr>

		<TD valign="top" colspan="2">

												<table id="tabSale" class="table" cellspacing="0" cellpadding="0" width="100%" align="center" border="0">

			<tr>

				<TD colspan="2"><table cellspacing="0" rules="all" border="1" id="dgGoods" width="100%">

					<tr>

						<td align="center" valign="middle" width="10%">门店代码</td><td align="center" valign="middle" width="45%">门店名称</td><td align="center" valign="middle" width="10%">销售日期</td><td align="center" valign="middle" width="15%">批号</td><td align="center" valign="middle" width="10%">数量</td><td align="center" valign="middle" width="10%">供应价</td>

					</tr><tr>

						<td align="center">103</td><td align="left">张江店</td><td align="center">2014年6月5日</td><td align="center">13092401</td><td align="right">6.0000</td><td align="right">0.000000000000</td>

					</tr><tr>

						<td align="center">108</td><td align="left">高桥店</td><td align="center">2014年6月16日</td><td align="center">13092401</td><td align="right">2.0000</td><td align="right">0.000000000000</td>

					</tr><tr>

						<td align="center">218</td><td align="left">乳山店</td><td align="center">2014年6月5日</td><td align="center">13092401</td><td align="right">11.0000</td><td align="right">0.000000000000</td>

					</tr><tr>

						<td align="center">219</td><td align="left">福山店</td><td align="center">2014年5月14日</td><td align="center">13092401</td><td align="right">2.0000</td><td align="right">0.000000000000</td>

					</tr><tr>

						<td align="center">311</td><td align="left">合庆店</td><td align="center">2014年6月17日</td><td align="center">13092401</td><td align="right">5.0000</td><td align="right">0.000000000000</td>

					</tr><tr>

						<td align="center">408</td><td align="left">凌桥店</td><td align="center">2014年6月4日</td><td align="center">13092401</td><td align="right">2.0000</td><td align="right">0.000000000000</td>

					</tr><tr>

						<td align="center">428</td><td align="left">光明路店</td><td align="center">2014年6月18日</td><td align="center">13092401</td><td align="right">2.0000</td><td align="right">0.000000000000</td>

					</tr><tr>

						<td align="center">合计</td><td align="left">&nbsp;</td><td align="center">&nbsp;</td><td align="center">&nbsp;</td><td align="right">30.0000</td><td align="right">0.000000000000</td>

					</tr><tr>

						<td align="center">现库存</td><td align="left">&nbsp;</td><td align="center">&nbsp;</td><td align="center">&nbsp;</td><td align="right">15.0000</td><td align="right">0.000000000000</td>

					</tr>

				</table>

															

														</TD>

			</tr>

		</table>

		

											</TD>

	</tr>

</table>



								</TD>

							</TR>

						</TABLE>

					</TD>

				</TR>

				<TR>

					<td vAlign="bottom" align="center">

版权所有 2005上海养和堂药业连锁经营有限公司

</td>

				</TR>

			</TABLE>

		</form>

	</body>

</HTML>
