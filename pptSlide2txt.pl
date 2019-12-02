use strict;
use warnings;
use Data::Dumper;
use utf8;
use HTML::Entities;
use XML::Twig;
use File::Find::Rule;
use File::Basename qw/basename dirname fileparse/;
use FindBin;
use File::Copy;
use Cwd;
use Archive::Zip qw( :ERROR_CODES :CONSTANTS :MISC_CONSTANTS );
use File::Path 'rmtree';;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft PowerPoint';

$Win32::OLE::Warn = 3;

binmode STDIN,  ':encoding(cp932)';
binmode STDOUT, ':encoding(cp932)';
binmode STDERR, ':encoding(cp932)';

#########################################################################
#                                                                      
# pptおよびpptxのスライドからテキストを抽出
# 
#########################################################################

# コマンドライン引数
my $arg = shift;
if (defined $arg){
	print "Option: " . $arg . "\n";
	# -nfn以外の引数だった場合は終了
	if ($arg ne '-nfn'){
		print "ERROR: Argument option accept 'nfn' only." . "\n";
		exit;
	}
}

#Local time settings
my $times = time();
my ($sec, $min, $hour, $mday, $month, $year, $wday, $stime) = localtime($times);
$month++;
my $datetime = sprintf '%04d%02d%02d%02d%02d%02d', $year + 1900, $month, $mday, $hour, $min, $sec;

# スクリプトファイルのディレクトリ
my $script_dir = $FindBin::Bin;

# tmpフォルダがなかったら作る
if ( -d './tmp' ){

} else {
	mkdir 'tmp', 0700 or die "$!";
}

print "Folder: ";
chomp( my $dir = <STDIN> );
$dir =~ s{^"}{};
$dir =~ s{"$}{};
my @pptxs = File::Find::Rule->file->name( '*.ppt', '*.pptx' )->in($dir);

# resultファイルをスクリプトのあるディレクトリに作る
chdir $script_dir;
open( my $out, ">:utf8", "result_$datetime.txt" ) or die "$!:result_$datetime.txt";

my @text;

foreach (@pptxs){
	my $pptx_fullpath = $_;
	my $pptx_filename = basename($pptx_fullpath);
	my $ppt_filename_not_ex = $pptx_filename;
	$ppt_filename_not_ex =~ s/^(.+)\.ppt$/$1/;
	my $pptx_dirname  = dirname($pptx_fullpath);
	print "\n" . "Processing... " . $pptx_filename . "\n";

	# 引数オプションが-nfnだったらファイル名区切りは出力しない。
	if ( defined $arg ){

	} else {
		push (@text, "\n\n------------------------------$pptx_filename------------------------------");
	}
	
	# pptとpptxをtmpフォルダに移動してzipにする。その際、pptはpptxに変換する。返り値はzipのファイル名。
	my $zip = &pptGenaretePptx_andCopyPptAndPptx2tmp($pptx_fullpath, $pptx_filename, $ppt_filename_not_ex, $pptx_dirname);

	chdir "$script_dir/tmp";

	# zip解凍
	&unzip(\$zip);

	# 展開後のzipを削除
	unlink $zip;

	my @xmls = File::Find::Rule->file->name( qr/slide\d+\.xml$/ )->in(getcwd);

	# ファイル名のスライド数を4桁に揃え、対象ファイルをtmpフォルダにコピーする
	# ※ファイル名が「slide1.xml」「slide10.xml」のようになっているので、処理順序を揃えるため。10000以上のスライドは想定外。
	&xml_rename_and_copy(\@xmls);

	# 要らないフォルダを削除
	&del_dir();
	
	# 4桁にリネームしてコピーしたxmlを対象とする
	my @xmls_rename = File::Find::Rule->file->name( qr/slide\d+\.xml$/ )->in(getcwd);
	
	# xmlをパースしてテキストをゲットする
	&xml_parser(\@xmls_rename);
	
	# 対象ファイルの*.xmlを削除する
	unlink glob '*.xml';
}

# resultファイルへの書き出し
for ( my $i=0; $i < scalar(@text); $i++ ) {
	my $text_out = $text[$i];
	print {$out} $text_out."\n";
}

close($out);

print "\n\n".'Done!'."\n";

# resultファイルをテキストエディタで開く
chdir $script_dir;
my $result = `result_$datetime.txt`;


sub pptGenaretePptx_andCopyPptAndPptx2tmp {
	my ($pptx_fullpath, $pptx_filename, $ppt_filename_not_ex, $pptx_dirname) = @_;
	my $zip;

	if ( $pptx_filename =~ /.+\.ppt$/ ){
		
		my $ppt = Win32::OLE->new('PowerPoint.Application','Quit');
		$ppt->{Visible} = 0; # 0で非表示にすると例外が発生することがある
		my $ppt_filehandle = $ppt->Presentations->Open($pptx_fullpath) or die $!;
		$pptx_filename = $ppt_filename_not_ex . '.pptx';
		$ppt_filehandle->SaveAs($pptx_dirname . '/' . $pptx_filename);
		$ppt_filehandle->Close();

		$zip = $pptx_filename;
		$zip =~ s|^(.+)$|$1\.zip|;

		copy($pptx_dirname . '/' . $pptx_filename, "$script_dir/tmp/$zip") or die $!;
		unlink $pptx_dirname . '/' . $pptx_filename;
		return $zip;
	} else {
		$zip = $pptx_filename;
		$zip =~ s|^(.+)$|$1\.zip|;
		copy($pptx_fullpath, "$script_dir/tmp/$zip") or die $!;
		return $zip;
	}
}

sub unzip {
	my ($zip) = shift;
	my $zip_obj = Archive::Zip->new($$zip);
	my @zip_members = $zip_obj->memberNames();
	foreach (@zip_members) {
		$zip_obj->extractMember($_, "$script_dir/tmp/$_");
	}
}

sub xml_rename_and_copy {
	my ($xmls) = shift;
	foreach (@$xmls){
		my $file_src = $_;
		my $file_dst;
		if ( $file_src =~ m|^(.+?/)(slide[0-9]\.xml)$| ){ # 1桁だったら
			$file_dst = $file_src;
			$file_dst =~ s|^(.+?/)(slide)([0-9])(\.xml)$|${2}000$3$4|;
			copy($file_src, $file_dst) or die {$!};
		} elsif ( $file_src =~ m|^(.+?/)(slide[0-9]{2,2}\.xml)$| ){ # 2桁だったら
			$file_dst = $file_src;
			$file_dst =~ s|^(.+?/)(slide)([0-9]{2,2})(\.xml)$|${2}00$3$4|;
			copy($file_src, $file_dst) or die {$!};
		} elsif ( $file_src =~ m|^(.+?/)(slide[0-9]{3,3}\.xml)$| ){ # 3桁だったら
			$file_dst = $file_src;
			$file_dst =~ s|^(.+?/)(slide)([0-9]{3,3})(\.xml)$|${2}0$3$4|;
			copy($file_src, $file_dst) or die {$!};
		}
		else {
			print "Error: The number of slides exceeds 1000.";
			exit;
		}
	}
}

sub del_dir {
	rmtree("$script_dir/tmp/ppt") or die $!;
	rmtree("$script_dir/tmp/docProps") or die $!;
	rmtree("$script_dir/tmp/_rels") or die $!;
}

sub xml_parser {
	my ($xmls_rename) = shift;
	foreach my $xml ( @$xmls_rename ){
		print $xml . "\n";
		my $twig = new XML::Twig( TwigRoots => {
				'//p:txBody|//a:txBody/a:p' => \&output_target,
				});
		$twig->parsefile( $xml );
	}
}

sub output_target {
	my( $tree, $elem ) = @_;
	my $target = $elem->text;
	push (@text, $target);
	
	{
		local *STDOUT;
		local *STDERR;
  		open STDOUT, '>', undef;
  		open STDERR, '>', undef;
		$tree->flush_up_to( $elem ); #Memory clear
	}
}
