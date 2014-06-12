use Win32::Console;
use Encode;
use 5.010;
use AnyEvent;

 
use Win32::Console::ANSI;
use Term::ANSIColor;
my @color    = qw( red  green  yellow  blue  magenta  cyan  white 
                   bright_black  bright_red  bright_green  bright_yellow 
				   bright_blue   bright_magenta  bright_cyan   bright_white ansi0);


#my @color= map {'ansi'.$_} (0..15);
$|=1; #必须开启这个
system("mode con cols=135 lines=25");
my $Out = new Win32::Console(STD_OUTPUT_HANDLE) || die;
my $cv = AnyEvent->condvar;
my $count=0;
my $w; $w = AnyEvent->timer(
        after       => 2, 
        interval => 2,
        cb => sub {
            $count++;
			
my ( $x, $y ) = $Out->Cursor();
$Out->Cursor( $x+125, $y + 5,0,0);

           while (<DATA>) {
  s/ /　/g;
  chomp;
  $a=decode('gb2312',$_);
  @words=$a=~m/(.)/g;
  

  foreach $word (@words) {  
         $c = $color[int rand @color];
		 print color 'bold '.$c;
         $Out->Write(encode('gb2312',$word));
         my ( $x, $y ) = $Out->Cursor();
         $Out->Cursor( $x, $y + 1,0,1);

         my ( $x, $y ) = $Out->Cursor();
         $Out->Cursor( $x-2, $y,0,1);
         select(undef,undef,undef,0.045);
         }

my ( $x, $y ) = $Out->Cursor();
$Out->Cursor( $x-2, $y-@words,0,0);
}
			system("cls");
			# say $count;
			seek(DATA,1565,0); # 句柄读到最后了就没有文本了，所以需要返回，其实DATA就是整个文件，是这个script的所有文本。
	
            if ($count >= 10) {
                undef $w; 
            }   
        }   
        );  
$cv->recv;
__END__




　　　　星月神话
我的一生最美好的场景
　
就是遇见你
　
在人海茫茫中静静凝望着你
　
陌生又熟悉
　
尽管呼吸着同一天空的气息

却无法拥抱到你
　
如果转换了时空身份和姓名
　　　
但愿认得你眼睛
　
千年之后的你会在哪里
　
身边有怎样风景
　
我们的故事并不算美丽
　
却如此难以忘记
　 
尽管呼吸着同一天空的气息
　
却无法拥抱到你
　 
如果转换了时空身份和姓名
　
但愿认得你眼睛
　
千年之后的你会在哪里
　
身边有怎样风景
　
我们的故事并不算美丽
　
却如此难以忘记
　
如果当初勇敢的在一起
　 
会不会不同结局
　
你会不会也有千言万语
　
埋在沉默的梦里　