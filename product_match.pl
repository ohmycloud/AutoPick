use strict 'vars';  # 声明全局变量不会报错
use warnings;

# 补全到 61000000
use List::MoreUtils qw/all/;
use Spreadsheet::XLSX;  
use Excel::Writer::XLSX;
use MyExcelFormatter;
use Encode;
use v5.10;
use Data::Dumper;

my %match;
while(<DATA>){
chomp;
my ($product_code,$product_name)=split(/\s+/,$_,2);
# say $product_code;
$product_name=~s{/}{}g;
# say $product_name;
$match{$product_code}=$product_name;
}

sub H{
my $text = shift;
return  decode('utf8',$text);  # 进行转码
}

my $workbook       =  Excel::Writer::XLSX->new( "模板.xlsx" );
my $excel          =  Spreadsheet::XLSX -> new ('产品匹配维护.xlsx');
my $n=2;
our $encode_name; # 声明一个全局变量

sub is{
all {$encode_name =~ /$_/i} @_;
}


my @ID_5300000 = qw( 百多邦 喷雾 70ml                       );
my @ID_5300000_2 = qw( 百多邦 70ml                          );
my @ID_5000000 = qw( 莫匹罗星 10g                           );
my @ID_5000000_2 = qw( 百多邦 10g                           );
my @ID_02251   = qw( 莫匹罗星 5g                            );
my @ID_02251_2 = qw( 莫匹罗星软膏                           );
my @ID_02251_3 = qw( 莫匹罗星  5克                          );
my @ID_02251_4 = qw( 百多邦膏  5克                          );
my @ID_5100000 = qw( 丁酸氯倍他松乳膏                       );
my @ID_9090012 = qw( 保丽净 假牙 局部 赠 24T                );
my @ID_9090013 = qw( 保丽净 假牙 赠 25g                     );
my @ID_9020001 = qw( 保丽净 假牙 6片                        );
my @ID_9090005 = qw( 保丽净 新年 促销 48T                   );
my @ID_9090003 = qw( 保丽净 促销装 假牙清洁盒 24T           );
my @ID_9090019 = qw( 保丽净 假牙 清洁 6片 促销              );
my @ID_9090021 = qw( 保丽净 假牙 清洁 60片 赠 6片 促销装    );
my @ID_9090023 = qw( 保丽净 局部 假牙 24片 赠购物袋         );
my @ID_9090201 = qw( 保丽净 假牙 清洁片 24片 赠放大镜       );
my @ID_9090202 = qw( 保丽净 局部 24片 赠手电筒              );
my @ID_9020405 = qw( 保丽净 假牙 黏合剂 60g 台湾再包装      );
my @ID_9020404 = qw( 保丽净 假牙 黏合剂 40g 台湾再包装      );
my @ID_9020401 = qw( 保丽净 假牙 稳固剂 70g                 );
my @ID_9020102 = qw( 保丽净 假牙 清洁片 30片 药店专供       );
my @ID_9020102_2 = qw( 保丽净 全 半口 30                    );
my @ID_9020102_3 = qw( 保丽净 全 半口 30 24 盒              );
my @ID_9020302 = qw( 假牙清洁片 局部 30                     );
my @ID_9020101 = qw( 假牙清洁片 24                          );
my @ID_9020103 = qw( 保丽净 假牙 清洁片 60                  );
my @ID_9020301 = qw( 假牙清洁片 局部  24                    );
my @ID_9090205 = qw( 保丽净 假牙 清洁片 24片 赠6片          );
my @ID_9090206 = qw( 保丽净局部24片赠6片                    );
my @ID_9090207 = qw( 保丽净假牙清洁片60片赠24片             );
my @ID_26151   = qw( 必理通 对乙酰氨基酚片 10               );
my @ID_26151_2 = qw( 对乙酰氨基酚片                         );
my @ID_44351   = qw( 肠虫清                                 );
my @ID_44351_2 = qw( 阿苯达唑片                             );
my @ID_10152   = qw( 布洛芬缓释胶囊 0.3 20                  );
my @ID_10152_2 = qw( 布洛芬缓释胶囊 300 20                  );
my @ID_10152_3 = qw( 芬必得 0.3 20                          );
my @ID_10152_4 = qw( 芬必得 胶囊 20                         );
my @ID_10152_5 = qw( 布洛芬 胶囊 20                         );
my @ID_5200000 = qw( 芬必得0.3g*4T                          );
my @ID_1035100 = qw( 布洛芬缓释胶囊 0.4                     );
my @ID_1035100_2 = qw( 芬必得胶囊 0.4                       );
my @ID_1035100_3 = qw( 布洛芬缓释胶囊 0.4 24                );
my @ID_1035100_4 = qw( 布洛芬缓释胶囊 400                   );
my @ID_1035100_5 = qw( 芬必得 400                           );
my @ID_26352   = qw(  酚咖片 20                             );
my @ID_26352_1 = qw(  芬咖片 20                             );
my @ID_26352_2 = qw(  酚咖片 10 2板                         );
my @ID_26352_3 = qw(  酚咖片 10 2B                          );
my @ID_26351   = qw(  酚咖片 10                             );
my @ID_10251   = qw( 布洛芬乳膏  20                         );
my @ID_10251_2 = qw( 芬必得乳膏  20                         );
my @ID_10251_3 = qw( 芬必得乳膏                             );
my @ID_10251_4 = qw( 芬必得乳膏                             );
my @ID_51002   = qw( 辅舒良 120                             );
my @ID_51002_2 = qw( 丙酸氟替卡松鼻喷雾剂 120               );
my @ID_135002G = qw( 施泰福 艾丽婷 修润沐浴油 150ml         );
my @ID_135003G = qw( 施泰福 乐蒂克 果酸保湿乳液 100g        );
my @ID_135004G = qw( 施泰福 爱可妮 洁肤露 100ml             );
my @ID_135005G = qw( 霏丝佳修润洁肤露100ml                  );
my @ID_135006G = qw( 施泰福 爱可妮 洁肤皂 100g              );
my @ID_135011G = qw( 施泰福 诗蓓白柔皙防晒乳霜 SPF30/PA 60g );
my @ID_135016G = qw( 施泰福 霏丝佳 润肤霜 75ml              );
my @ID_135020G = qw( 施泰福 霏丝佳 润肤乳液 100ml           );
my @ID_135021G = qw( 霏丝佳 特护修润霜 50ml                 );
my @ID_135033G = qw( 霏丝佳 特护修润乳液 100ml              );
my @ID_135034G = qw( 施泰福 霏丝佳 修润密集滋养霜 50ml      );
my @ID_135035G = qw( 施泰福 霏丝佳 修润沐浴露 150ml         );
my @ID_9090017 = qw( 舒适达 牙龈护理 120g 赠 速效 25g       );
my @ID_9090018 = qw( 舒适达 专业修复 100g 赠 马克杯         );
my @ID_9090016 = qw( 舒适达 速效抗敏 120g 赠 速效 25        );
my @ID_9090009 = qw( 舒适达 速效抗敏 120g 送牙刷            );
my @ID_9090011 = qw( 舒适达 美白 120g 赠牙刷                );
my @ID_9090010 = qw( 舒适达 美白 120g 赠马克杯              );
my @ID_9090008 = qw( 舒适达 清新 120g 赠 清新 25g 促销装    );
my @ID_9090004 = qw( 舒适达 清新 120g 赠 速效 25g 促销装    );
my @ID_9090002 = qw( 舒适达 升级 全面护理 50g 特价体验装    );
my @ID_9010005 = qw( 舒适达 美白 120                        );
my @ID_9010002 = qw( 舒适达  清新 薄荷 120                  );
my @ID_9010006 = qw( 舒适达 全效护理 50g                    );
my @ID_9010012 = qw( 舒适达 全面护理 漱口水 500ml           );
my @ID_9010013 = qw( 舒适达 全面护理 漱口水 再包装 500ml    );
my @ID_9010401 = qw( 舒适达 牙龈护理 180g                   );
my @ID_9010302 = qw( 舒适达 速效抗敏 180g                   );
my @ID_9090102 = qw( 舒适达 全面护理 120g 赠牙刷            );
my @ID_9010301 = qw( 舒适达 速效抗敏 100g                   );
my @ID_9010009 = qw( 舒适达 清新薄荷 25g                    );
my @ID_9090001 = qw( 舒适达 清新薄荷 双支 88折优惠装        );
my @ID_9090103 = qw( 舒适达 牙龈护理 120g 赠牙刷 促销装     );
my @ID_9090104 = qw( 舒适达 牙龈护理 120克 赠全面护理 50g   );
my @ID_9010701 = qw( 舒适达 微粒劲洁泡沫啫喱牙膏  100ml     );
my @ID_9010017 = qw( 舒适达 全面护理  120                   );
my @ID_9010018 = qw( 舒适达 速效抗敏      120               );
my @ID_9010019 = qw( 舒适达 牙龈护理      120               );
my @ID_9090105 = qw( 舒适达 专业修复 100g 赠 25             );
my @ID_9090106 = qw( 舒适达 专业修复 美白 100g 赠 25g       );
my @ID_9090107 = qw( 舒适达 专业修复 100g 赠乐扣杯          );
my @ID_9090108 = qw( 舒适达 专业修复 美白 赠乐扣杯          );
my @ID_9010101 = qw( 全面护理 70                            );
my @ID_9010303 = qw( 速效抗敏 70                            );
my @ID_9090109 = qw( 舒适达 专业修复100克赠速效25g(OTC专供) );
my @ID_9010800 = qw( 舒适达 专业修复 100                    );
my @ID_9011400 = qw( 舒适达 专业修复 美白 100               );
my @ID_9010900 = qw( 舒适达 全方位防护 100                  );
my @ID_9011000 = qw( 舒适达 全方位防护 劲爽 薄荷 100        );
my @ID_9090114 = qw( 舒适达 速效抗敏 赠 50                  );
my @ID_9090110 = qw( 舒适达 全面护理 赠 25g                 );
my @ID_9010391 = qw( 舒适达 速效抗敏 25                     );
my @ID_9010990 = qw( 舒适达 全方位防护 27                   );
my @ID_9010390 = qw( 舒适达 速效抗敏 25                     );
my @ID_9010890 = qw( 舒适达 专业修复 27                     );
my @ID_9090115 = qw( 舒适达 速效抗敏 180  赠 全面护理 50    );
my @ID_9010102 = qw( 舒适达 全面护理 180                    );
my @ID_9010501 = qw( 舒适达 美白配方 180                    );
my @ID_9090116 = qw( 舒适达 全面护理 120g 赠 25g            );
my @ID_6100000 = qw( 通气鼻贴 透明                          );
my @ID_6100000_2 = qw( 通气鼻帖 透明                        );
my @ID_6100100 = qw( 通气鼻贴 肤色                          );
my @ID_6100100_2 = qw(通气鼻帖 肤色                         );
my @ID_6100200 = qw( 通气鼻贴 儿童  8                       );
my @ID_6100201 = qw( 通气鼻帖 儿童  促销装                  );
my @ID_6100400 = qw( 通气鼻贴 薄荷                          );
my @ID_7200100 = qw(  软糖 柠檬 20  袋                      ); 
my @ID_7200000 = qw( 软糖 柠檬 20                           );
my @ID_7200000_2 = qw( 软糖 柠檬 40                         );
my @ID_7100100 = qw( 软糖 薄荷 20 袋                        );
my @ID_7110000 = qw(  软糖 薄荷 20 G                        );
my @ID_7110000_2 = qw(  软糖 薄荷 40                        );
my @ID_7300000 = qw(  软糖 莓果                             );
my @ID_7300000_2 = qw(  新康泰克 莓果                       );
my @ID_7300000_3 = qw(  新康泰克 草莓                       );
my @ID_7300000_4 = qw(  软糖 草莓                           );
my @ID_7100002 = qw( 喉爽 薄荷 听装 促销装                  );
my @ID_7200002 = qw( 软糖 柠檬 40G                       );
my @ID_7300003 = qw( 润喉软糖莓果口味+薄荷袋装促销装 听 袋  );
my @ID_05451   = qw( 新康泰克胶囊 10                        );
my @ID_05451_2 = qw( 复方盐酸伪麻黄  10                     );
my @ID_0545102 = qw( 盐酸伪麻黄  8                          );
my @ID_0545102_2 = qw( 新康泰克 胶囊 8                      );
my @ID_05252   = qw( 美扑伪麻 10                            );
my @ID_05252_2 = qw( 美扑伪麻 10 200                        );
my @ID_05252_3 = qw( 美扑伪麻片 10S 10盒 20条               );
my @ID_05253   = qw( 美扑伪麻 20                            );
my @ID_05253_2 = qw( 美扑伪麻 10 2板                        );
my @ID_6600100 = qw( 盐酸氨溴索缓释胶囊 10                  );
my @ID_6600000 = qw( 盐酸氨溴索缓释胶囊 6                   );
my @ID_9030201 = qw( 益周适 专业牙龈 护理牙膏 劲爽薄荷 180g );
my @ID_9030100 = qw( 益周适 专业牙龈  120g                  );
my @ID_9030100_2 = qw( 益周适 专业牙龈  120克               );
my @ID_9030101 = qw( 益周适 专业牙龈 护理牙膏 180g          );
my @ID_9030200 = qw( 益周适 专业牙龈 护理牙膏 劲爽薄荷 120g );
my @ID_knl     = qw( 康纳乐                                 );
my @ID_knl_2   = qw( 复方曲安奈德乳膏                       );
my @ID_bkl     = qw( 伯克纳                                 );
my @ID_bkl_2   = qw( 丙酸倍氯米松鼻喷雾剂                   );
my @ID_hpd     = qw( 贺普丁                                 );
my @ID_hpd_2   = qw( 拉米夫定片                             );
my @ID_hwl     = qw( 贺维力                                 );
my @ID_hwl_2   = qw( 阿德福韦酯片                           );
my @ID_lms     = qw( 兰美抒                                 );
my @ID_lms_2   = qw( 盐酸特比萘芬乳膏                       );
my @ID_slt     = qw( 赛乐特                                 );
my @ID_slt_2   = qw( 盐酸帕罗西汀片                         );
my @ID_twm     = qw( 泰为美                                 );
my @ID_twm_2   = qw( 西咪替丁片                             );
my @ID_wtl     = qw( 万托林                                 );
my @ID_wtl_2   = qw( 硫酸沙丁胺醇吸入气雾剂                 );
my @ID_smt     = qw( 沙美特                                 );
my @ID_smt_2   = qw( 罗体卡松粉                             );
my @ID_fst     = qw( 辅舒酮                                 );
my @ID_9090209 = qw(保丽净 30 送 6                          );
my @ID_9090208 = qw(保丽净  局部 30 送 6                    );
my @ID_9090209_2 = qw(保丽净 30  6                          );
my @ID_9090208_2 = qw(保丽净  局部 30  6                    );
# 7100003、7200003  、qita_fushuliang 这个编码不能用，为禁码 
our $a;                  
{
no warnings;
sub match{
given( $encode_name ) {
when( is( @ID_05451    ) ) {  $a =  ['05451',  '盒'];continue; }
when( is( @ID_05451_2  ) ) {  $a =  ['05451',  '盒'];continue; }
when( is( @ID_5300000  ) ) {  $a =  ['5300000','瓶'];continue; }
when( is( @ID_5300000_2) ) {  $a =  ['5300000','瓶'];continue; }
when( is( @ID_5000000  ) ) {  $a =  ['5000000','支'];continue; }
when( is( @ID_5000000_2) ) {  $a =  ['5000000','支'];continue; }
when( is( @ID_02251    ) ) {  $a =  ['02251',  '支'];continue; }
when( is( @ID_02251_2  ) ) {  $a =  ['02251',  '支'];continue; }
when( is( @ID_02251_3  ) ) {  $a =  ['02251',  '支'];continue; }
when( is( @ID_02251_4  ) ) {  $a =  ['02251',  '支'];continue; }
when( is( @ID_5100000  ) ) {  $a =  ['5100000','盒'];continue; }
when( is( @ID_9090012  ) ) {  $a =  ['9090012','盒'];continue; }
when( is( @ID_9090013  ) ) {  $a =  ['9090013','盒'];continue; }
when( is( @ID_9020001  ) ) {  $a =  ['9020001','盒'];continue; }
when( is( @ID_9090005  ) ) {  $a =  ['9090005','盒'];continue; }
when( is( @ID_9090003  ) ) {  $a =  ['9090003','盒'];continue; }
when( is( @ID_9090019  ) ) {  $a =  ['9090019','盒'];continue; }
when( is( @ID_9090021  ) ) {  $a =  ['9090021','盒'];continue; }
when( is( @ID_9090023  ) ) {  $a =  ['9090023','盒'];continue; }
when( is( @ID_9090201  ) ) {  $a =  ['9090201','盒'];continue; }
when( is( @ID_9090202  ) ) {  $a =  ['9090202','盒'];continue; }
when( is( @ID_9020405  ) ) {  $a =  ['9020405','支'];continue; }
when( is( @ID_9020404  ) ) {  $a =  ['9020404','支'];continue; }
when( is( @ID_9020401  ) ) {  $a =  ['9020401','支'];continue; }
when( is( @ID_9020101  ) ) {  $a =  ['9020101','盒'];continue; }
when( is( @ID_9020102  ) ) {  $a =  ['9020102','盒'];continue; }
when( is( @ID_9020102_2) ) {  $a =  ['9020102','盒'];continue; }
when( is( @ID_9020102_3) ) {  $a =  ['9020102','盒'];continue; }
when( is( @ID_9020103  ) ) {  $a =  ['9020103','盒'];continue; }
when( is( @ID_9020301  ) ) {  $a =  ['9020301','盒'];continue; }
when( is( @ID_9020302  ) ) {  $a =  ['9020302','盒'];continue; }
when( is( @ID_9090205  ) ) {  $a =  ['9090205','盒'];continue; }
when( is( @ID_9090206  ) ) {  $a =  ['9090206','盒'];continue; }
when( is( @ID_9090207  ) ) {  $a =  ['9090207','盒'];continue; }
when( is( @ID_26151    ) ) {  $a =  ['26151',  '盒'];continue; }
when( is( @ID_26151_2  ) ) {  $a =  ['26151',  '盒'];continue; }
when( is( @ID_44351    ) ) {  $a =  ['44351',  '盒'];continue; }
when( is( @ID_44351_2  ) ) {  $a =  ['44351',  '盒'];continue; }
when( is( @ID_10152    ) ) {  $a =  ['10152',  '盒'];continue; }
when( is( @ID_10152_2  ) ) {  $a =  ['10152',  '盒'];continue; }
when( is( @ID_10152_3  ) ) {  $a =  ['10152',  '盒'];continue; }
when( is( @ID_10152_4  ) ) {  $a =  ['10152',  '盒'];continue; }
when( is( @ID_10152_5  ) ) {  $a =  ['10152',  '盒'];continue; }
when( is( @ID_5200000  ) ) {  $a =  ['5200000','盒'];continue; }
when( is( @ID_1035100  ) ) {  $a =  ['1035100','盒'];continue; }
when( is( @ID_1035100_2) ) {  $a =  ['1035100','盒'];continue; }
when( is( @ID_1035100_3) ) {  $a =  ['1035100','盒'];continue; }
when( is( @ID_1035100_4) ) {  $a =  ['1035100','盒'];continue; }
when( is( @ID_1035100_5) ) {  $a =  ['1035100','盒'];continue; }
when( is( @ID_26351    ) ) {  $a =  ['26351',  '盒'];continue; }
when( is( @ID_26352    ) ) {  $a =  ['26352',  '盒'];continue; }
when( is( @ID_26352_1  ) ) {  $a =  ['26352',  '盒'];continue; }
when( is( @ID_26352_2  ) ) {  $a =  ['26352',  '盒'];continue; }
when( is( @ID_26352_3  ) ) {  $a =  ['26352',  '盒'];continue; }
when( is( @ID_10251    ) ) {  $a =  ['10251',  '支'];continue; }
when( is( @ID_10251_2  ) ) {  $a =  ['10251',  '支'];continue; }
when( is( @ID_10251_3  ) ) {  $a =  ['10251',  '支'];continue; }
when( is( @ID_10251_4  ) ) {  $a =  ['10251',  '支'];continue; }
when( is( @ID_51002    ) ) {  $a =  ['51002',  '支'];continue; }
when( is( @ID_51002_2  ) ) {  $a =  ['51002',  '支'];continue; }
when( is( @ID_135002G  ) ) {  $a =  ['135002G','支'];continue; }
when( is( @ID_135003G  ) ) {  $a =  ['135003G','支'];continue; }
when( is( @ID_135004G  ) ) {  $a =  ['135004G','支'];continue; }
when( is( @ID_135005G  ) ) {  $a =  ['135005G','支'];continue; }
when( is( @ID_135006G  ) ) {  $a =  ['135006G','支'];continue; }
when( is( @ID_135011G  ) ) {  $a =  ['135011G','支'];continue; }
when( is( @ID_135016G  ) ) {  $a =  ['135016G','支'];continue; }
when( is( @ID_135020G  ) ) {  $a =  ['135020G','支'];continue; }
when( is( @ID_135021G  ) ) {  $a =  ['135021G','支'];continue; }
when( is( @ID_135033G  ) ) {  $a =  ['135033G','支'];continue; }
when( is( @ID_135034G  ) ) {  $a =  ['135034G','支'];continue; }
when( is( @ID_135035G  ) ) {  $a =  ['135035G','支'];continue; }
when( is( @ID_9090017  ) ) {  $a =  ['9090017','支'];continue; }
when( is( @ID_9090018  ) ) {  $a =  ['9090018','支'];continue; }
when( is( @ID_9090016  ) ) {  $a =  ['9090016','支'];continue; }
when( is( @ID_9090009  ) ) {  $a =  ['9090009','支'];continue; }
when( is( @ID_9090011  ) ) {  $a =  ['9090011','支'];continue; }
when( is( @ID_9090010  ) ) {  $a =  ['9090010','支'];continue; }
when( is( @ID_9090008  ) ) {  $a =  ['9090008','支'];continue; }
when( is( @ID_9090004  ) ) {  $a =  ['9090004','支'];continue; }
when( is( @ID_9090002  ) ) {  $a =  ['9090002','支'];continue; }
when( is( @ID_9010005  ) ) {  $a =  ['9010005','支'];continue; }
when( is( @ID_9010002  ) ) {  $a =  ['9010002','支'];continue; }
when( is( @ID_9010006  ) ) {  $a =  ['9010006','支'];continue; }
when( is( @ID_9010012  ) ) {  $a =  ['9010012','支'];continue; }
when( is( @ID_9010013  ) ) {  $a =  ['9010013','支'];continue; }
when( is( @ID_9010401  ) ) {  $a =  ['9010401','支'];continue; }
when( is( @ID_9010302  ) ) {  $a =  ['9010302','支'];continue; }
when( is( @ID_9090102  ) ) {  $a =  ['9090102','盒'];continue; }
when( is( @ID_9010301  ) ) {  $a =  ['9010301','支'];continue; }
when( is( @ID_9010009  ) ) {  $a =  ['9010009','支'];continue; }
when( is( @ID_9090001  ) ) {  $a =  ['9090001','支'];continue; }
when( is( @ID_9090103  ) ) {  $a =  ['9090103','支'];continue; }
when( is( @ID_9090104  ) ) {  $a =  ['9090104','支'];continue; }
when( is( @ID_9010701  ) ) {  $a =  ['9010701','支'];continue; }
when( is( @ID_9010017  ) ) {  $a =  ['9010017','支'];continue; }
when( is( @ID_9010018  ) ) {  $a =  ['9010018','支'];continue; }
when( is( @ID_9010019  ) ) {  $a =  ['9010019','支'];continue; }
when( is( @ID_9090105  ) ) {  $a =  ['9090105','支'];continue; }
when( is( @ID_9090106  ) ) {  $a =  ['9090106','支'];continue; }
when( is( @ID_9090107  ) ) {  $a =  ['9090107','支'];continue; }
when( is( @ID_9090108  ) ) {  $a =  ['9090108','支'];continue; }
when( is( @ID_9010101  ) ) {  $a =  ['9010101','支'];continue; }
when( is( @ID_9010303  ) ) {  $a =  ['9010303','支'];continue; }
when( is( @ID_9090109  ) ) {  $a =  ['9090109','支'];continue; }
when( is( @ID_9010800  ) ) {  $a =  ['9010800','支'];continue; }
when( is( @ID_9011400  ) ) {  $a =  ['9011400','支'];continue; }
when( is( @ID_9010900  ) ) {  $a =  ['9010900','支'];continue; }
when( is( @ID_9011000  ) ) {  $a =  ['9011000','支'];continue; }
when( is( @ID_9090114  ) ) {  $a =  ['9090114','支'];continue; }
when( is( @ID_9090110  ) ) {  $a =  ['9090110','支'];continue; }
when( is( @ID_9010391  ) ) {  $a =  ['9010391','支'];continue; }
when( is( @ID_9010990  ) ) {  $a =  ['9010990','支'];continue; }
when( is( @ID_9010390  ) ) {  $a =  ['9010390','支'];continue; }
when( is( @ID_9010890  ) ) {  $a =  ['9010890','支'];continue; }
when( is( @ID_9090115  ) ) {  $a =  ['9090115','支'];continue; }
when( is( @ID_9010102  ) ) {  $a =  ['9010102','支'];continue; }
when( is( @ID_9010501  ) ) {  $a =  ['9010501','支'];continue; }
when( is( @ID_9090116  ) ) {  $a =  ['9090116','支'];continue; }
when( is( @ID_6100000  ) ) {  $a =  ['6100000','盒'];continue; }
when( is( @ID_6100000_2) ) {  $a =  ['6100000','盒'];continue; }
when( is( @ID_6100100  ) ) {  $a =  ['6100100','盒'];continue; }
when( is( @ID_6100100_2) ) {  $a =  ['6100100','盒'];continue; }
when( is( @ID_6100200  ) ) {  $a =  ['6100200','盒'];continue; }
when( is( @ID_6100201  ) ) {  $a =  ['6100201','盒'];continue; }
when( is( @ID_6100400  ) ) {  $a =  ['6100400','盒'];continue; }
when( is( @ID_7200100  ) ) {  $a =  ['7200100','袋'];continue; }
when( is( @ID_7200000  ) ) {  $a =  ['7200000','瓶'];continue; }
when( is( @ID_7200000_2) ) {  $a =  ['7200000','瓶'];continue; }
when( is( @ID_7100100  ) ) {  $a =  ['7100100','袋'];continue; }
when( is( @ID_7110000  ) ) {  $a =  ['7110000','瓶'];continue; }
when( is( @ID_7110000_2) ) {  $a =  ['7110000','瓶'];continue; }
when( is( @ID_7300000  ) ) {  $a =  ['7300000','瓶'];continue; }
when( is( @ID_7300000_2) ) {  $a =  ['7300000','瓶'];continue; }
when( is( @ID_7300000_3) ) {  $a =  ['7300000','瓶'];continue; }
when( is( @ID_7300000_4) ) {  $a =  ['7300000','瓶'];continue; }
when( is( @ID_7100002  ) ) {  $a =  ['7100002','瓶'];continue; }
when( is( @ID_7200002  ) ) {  $a =  ['7200002','瓶'];continue; }
when( is( @ID_7300003  ) ) {  $a =  ['7300003','瓶'];continue; }
when( is( @ID_0545102  ) ) {  $a =  ['0545102','盒'];continue; }
when( is( @ID_0545102_2) ) {  $a =  ['0545102','盒'];continue; }
when( is( @ID_05252    ) ) {  $a =  ['05252',  '盒'];continue; }
when( is( @ID_05253    ) ) {  $a =  ['05253',  '盒'];continue; }
when( is( @ID_05252_2  ) ) {  $a =  ['05252',  '盒'];continue; } # 更精确的匹配位置要靠后
when( is( @ID_05253_2  ) ) {  $a =  ['05253',  '盒'];continue; }
when( is( @ID_05252_3  ) ) {  $a =  ['05252',  '盒'];continue; }
when( is( @ID_6600100  ) ) {  $a =  ['6600100','盒'];continue; }
when( is( @ID_6600000  ) ) {  $a =  ['6600000','盒'];continue; }
when( is( @ID_9030201  ) ) {  $a =  ['9030201','支'];continue; }
when( is( @ID_9030100  ) ) {  $a =  ['9030100','支'];continue; }
when( is( @ID_9030100_2) ) {  $a =  ['9030100','支'];continue; }
when( is( @ID_9030101  ) ) {  $a =  ['9030101','支'];continue; }
when( is( @ID_9030200  ) ) {  $a =  ['9030200','支'];continue; }
when( is( @ID_knl      ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_knl_2    ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_bkl      ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_bkl_2    ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_hpd      ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_hpd_2    ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_hwl      ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_hwl_2    ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_lms      ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_lms_2    ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_slt      ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_slt_2    ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_twm      ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_twm_2    ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_wtl      ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_wtl_2    ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_smt      ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_smt_2    ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_fst      ) ) {  $a =  ['qita',   '盒'];continue; }
when( is( @ID_9090209  ) ) {  $a =  ['9090209','盒'];continue; }
when( is( @ID_9090208  ) ) {  $a =  ['9090208','盒'];continue; }
when( is( @ID_9090209_2) ) {  $a =  ['9090209','盒'];continue; }
when( is( @ID_9090208_2) ) {  $a =  ['9090208','盒'];continue; }
}
}
}

my @headers=(
['ID', '组织名称', '经销商编码', '经销商名称', '客户端编码', '客户端名称', '经销商产品编码', '经销商产品条码', '经销商产品名称', '经销商产品单位', '产品标准编码', '产品标准名称', 'UPC', '产品类型', '匹配后单位', '创建日期', '匹配日期', '最近变更日期']
);

my @encoded_headers;
foreach (@{$headers[0]}) {
push @{$encoded_headers[0]},decode("gb2312",$_);
};

my $format1     = $workbook->add_format( align => 'left',  font => 'Arial', size => 8);
my $format2     = $workbook->add_format( align => 'left',  font => 'Arial', size => 8,num_format => '@');
my @columns;

foreach my $old_sheet( @{$excel -> {Worksheet}}[0] ) {  # 只读取第一张工作表，直链安装
        my @columns; 
        my $new_worksheet  =  $workbook->add_worksheet(H($old_sheet->{Name}));  # 将原来的工作表名添加到新的 Excel
		$new_worksheet->keep_leading_zeros(); # 保留前导0
		$new_worksheet->write_col( 'A1', \@encoded_headers,$format1);
	    $old_sheet -> {MaxCol} ||= $old_sheet -> {MinCol};  #|| 逻辑或，返回计算结果先为真的值，由左到右计算
        	   
# 读取原 Excel，然后写入 新 Excel        
foreach my $col (0..16) {
			    $old_sheet -> {MaxRow} ||= $old_sheet -> {MinRow}; 
                foreach my $row ($old_sheet -> {MinRow} ..  $old_sheet -> {MaxRow}) {  
					    push @{$columns[$col]},H($old_sheet -> {Cells} [$row] [$col]->{Val});
						# push @{$columns[$col]},$old_sheet -> {Cells} [$row] [$col]->{Val};  # 提取固定的几列 
						# say '-'.$old_sheet -> {Cells} [$row] [$col]->{Val};
                    }
# shift @{$columns[$col]};
}
	$new_worksheet->write_col( 'A2', \@{$columns[0]},$format2);
	$new_worksheet->write_col( 'E2', \@{$columns[4]},$format1);
	$new_worksheet->write_col( 'G2', \@{$columns[5]},$format1);
	$new_worksheet->write_col( 'I2', \@{$columns[7]},$format1);
	$new_worksheet->write_col( 'J2', \@{$columns[15]},$format1);
	
	# print Dumper(\@{$columns[7]});
	
	for my $name ( @{$columns[7]} ) {
        chomp $name;
		$encode_name = encode("gb2312",$name);  # 编码后才能让匹配识别
		{
		  no warnings;
		  &match;
		  $new_worksheet->write( 'K'.$n,$a->[0],$format2);
		  $new_worksheet->write( 'O'.$n++,decode("gb2312",$a->[1]),$format2);
		  undef $a;
		}
   }	
}
__DATA__
产品代码	产品名称
5300000	百多邦喷雾
5000000	百多邦10g
02251	百多邦5g
5100000	百多邦皮炎湿疹乳膏15g
9090012	保丽净假牙清洁片局部赠清洁刷 24T
9090013	保丽净假牙清洁片赠速效25g  24T+25g
9020001	保丽净假牙清洁片6片体验装
9090005	保丽净新年促销装48T
9090003	保丽净买赠促销装 - 赠假牙清洁盒24T
9090019	保丽净假牙清洁6片促销装
9090021	保丽净假牙清洁60片赠6片促销装
9090023	保丽净局部假牙24片赠购物袋
9090201	保丽净假牙清洁片24片赠放大镜
9090202	保丽净局部24片赠手电筒
9020405	保丽净假牙黏合剂60g台湾再包装
9020404	保丽净假牙黏合剂40g台湾再包装
9020401	保丽净假牙稳固剂70g
9020102	保丽净假牙清洁片30片（药店专供）
9020302	保丽净假牙清洁片局部假牙专用30片（药店专供）
9020101	保丽净假牙清洁片24片
9020103	保丽净假牙清洁片60片
9020301	保丽净假牙清洁片局部假牙专用24片
9090205	保丽净假牙清洁片24片赠6片
9090206	保丽净局部24片赠6片
9090207	保丽净假牙清洁片60片赠24片
26151	必理通0.5gx10T
44351	肠虫清0.2gx10T
10152	芬必得0.3gx20T
5200000	芬必得0.3g*4T
1035100	芬必得布洛芬缓释胶囊 400mg
26352	芬必得 酚咖片20T
26351	芬必得 酚咖片10T
10251	芬必得膏
51002	辅舒良
135002G	施泰福艾丽婷修润沐浴油150ml
135003G	施泰福乐蒂克果酸保湿乳液100g
135004G	施泰福爱可妮洁肤露100ml
135005G	霏丝佳修润洁肤露100ml
135006G	施泰福爱可妮洁肤皂100g
135011G	施泰福诗蓓白柔皙防晒乳霜SPF30/PA++60g
135016G	施泰福霏丝佳润肤霜75ml
135020G	施泰福霏丝佳润肤乳液100ml
135021G	霏丝佳特护修润霜50ml
135033G	霏丝佳特护修润乳液100ml
135034G	施泰福霏丝佳修润密集滋养霜50ml
135035G	施泰福霏丝佳修润沐浴露 150ml
9090017	舒适达牙龈护理120g赠速效25g促销装
9090018	舒适达专业修复100g赠马克杯
9090016	舒适达速效抗敏120g赠速效25
9090009	舒适达速效抗敏120g送牙刷
9090011	舒适达美白120g赠牙刷
9090010	舒适达美白120g赠马克杯
9090008	舒适达清新120g赠清新25g促销装
9090004	舒适达清新120g赠速效25g促销装
9090002	舒适达升级全面护理50g特价体验装
9010005	舒适达美白120g
9010002	舒适达升级清新薄荷120g
9010006	舒适达全效护理50g
9010012	舒适达全面护理漱口水 500ml
9010013	舒适达全面护理漱口水(再包装）500ml
9010401	舒适达牙龈护理180g
9010302	舒适达速效抗敏180g
9090102	舒适达全面护理120g赠牙刷
9010301	舒适达速效抗敏100g
9010009	舒适达-清新薄荷25g
9090001	舒适达清新薄荷双支88折优惠装
9090103	舒适达牙龈护理120g赠牙刷促销装
9090104	舒适达牙龈护理120克赠全面护理50g
9010701	舒适达微粒劲洁泡沫啫喱牙膏（再包装）100ml
9010017	舒适达-升级全面护理(国产)120g
9010018	舒适达-速效抗敏(国产)120g
9010019	舒适达-牙龈护理(国产)120g
9090105	舒适达专业修复100g赠速效25
9090106	舒适达专业修复美白100g赠速效25g
9090107	舒适达专业修复100g赠乐扣杯
9090108	舒适达专业修复美白赠乐扣杯
9010101	全面护理70g
9010303	速效抗敏70g
9090109	舒适达专业修复100克赠速效25g(OTC专供)
9010800	舒适达专业修复100克
9011400	舒适达专业修复美白100克
9010900	舒适达全方位防护100克
9011000	舒适达全方位防护劲爽薄荷100克
9090114	舒适达速效抗敏赠全面50克
9090110	舒适达全面护理赠速效25g
9010391	舒适达-速效抗敏25gDDT包装(国产)
9010990	舒适达-全方位防护27g(国产)
9010390	舒适达-速效抗敏25g(国产)
9010890	舒适达-专业修复27g(国产)
9090115	舒适达速效抗敏180克赠全面护理50克
9010102	舒适达全面护理180克
9010501	舒适达美白配方180克
9090116	舒适达全面护理120g赠速效抗敏25g
6100000	新康泰克通气鼻贴透明型（标准）
6100100	新康泰克通气鼻贴肤色型（标准）10片
6100200	新康泰克通气鼻贴儿童型 8片
6100201	新康泰克通气鼻帖儿童型(促销装)
6100400	新康泰克通气鼻贴薄荷型
7200100	新康泰克喉爽草本润喉软糖-柠檬口味20g袋装
7200000	新康泰克喉爽草本润喉软糖-柠檬口味20粒
7100100	新康泰克喉爽草本润喉软糖-薄荷口味20g袋装
7110000	新康泰克喉爽草本润喉软糖-薄荷口味20粒(国产)
7300000	新康泰克喉爽_草本润喉软糖40g_莓果口味
7100002	新康泰克喉爽薄荷听装促销装
7200002	新康泰克喉爽柠檬听装促销装
7200003	新康泰克润喉软糖柠檬口味促销装(听+袋)
7100003	新康泰克润喉软糖薄荷口味促销装(听+袋)
7300003	新康泰克润喉软糖莓果口味+薄荷袋装促销装(听+袋)
05451	新康泰克10c
0545102	新康泰克感冒胶囊8粒
05252	新康泰克重感装美扑伪麻10T
05253	新康泰克重感装美扑伪麻20T
6600100	新康泰克盐酸氨溴索缓释胶囊10粒
6600000	新康泰克盐酸氨溴索缓释胶囊6粒
9030201	益周适专业牙龈护理牙膏劲爽薄荷180g
9030100	益周适专业牙龈护理牙膏120g
9030101	益周适专业牙龈护理牙膏180g
9030200	益周适专业牙龈护理牙膏劲爽薄荷120g