#!/usr/local/bin/perl
#
# $Id: utils.pl,v 1.3 2019/04/01 10:24:46 keenanf Exp keenanf $
# $Log: utils.pl,v $
# Revision 1.3  2019/04/01 10:24:46  keenanf
# Added routine &delete_unlinked_prongs($_)
#
# Revision 1.2  2019/04/01 10:17:28  keenanf
# *** empty log message ***
#
# Revision 1.1  2013/07/09 08:28:36  keenanf
# Initial revision
#
# Revision 1.1  2006/02/27 08:52:01  keenanf
# Initial revision
#
#
use File::Compare;
use Cwd 'abs_path';
require "./restructure.pl";
# $_ = &hex_non_ascii($_) 
# &mycomm($comm, $DBG, $quiet)
# $_ = &clean_wd($_);
# $_ = &reduce_idmids($_)
# $bit = &get_tag_contents($bit, "TAGNAME")
# $bit = &get_tag_attval
# $bit = &get_attval$e, "ATTNAME")
# &open_debug_files;
# &close_debug_files;
# ($file, $lang, $phon, $wd) = &get_phon_info($bit);
# &numerical_hash_sort(\%HASH, \@ARR); # puts the keys from the hash into the array in numerical order
# $bit = &get_name_date
# $bit = &my_name_date
# $pwd = &get_pwd;
# $time = &gmt;
# $bit = &get_xyz_file
# $bit = &get_fileh
# $bit = &get_target_fname
# $bit = &change_ae_ents
# $bit = &ents2letter
# $bit = &get_dr_old
# $bit = &get_hex_h
# $bit = &destress
# $bit = &get_dps_entry_id
# $bit = &get_entry_id
# $bit = &get_h
# $bit = &get_dr
# $bit = &get_xml_h
# $bit = &get_norm_hdwd
# $bit = &get_hdwd
# $bit = &get_hdwd_BAK
# $bit = &get_accented_hdwd
# $bit = &get_dr_old2
# $bit = &get_h_hm
# $bit = &get_topg
# $bit = &get_posg
# $bit = &get_xh_xhm
# $bit = &get_sortkey
# $bit = &get_sortwd
# $bit = &def_prevletter
# $bit = &def_nextletter
# $bit = &get_entry_file
# $bit = &get_date
# $bit = &get_date2
# $bit = &get_date3
# $bit = &get_logfile
# $bit = &open_logfile
# $bit = &get_sound_dir
# $bit = &mkdir
# $bit = &tidywd
# $bit = &tidy_hexwd
# $bit = &add_space_between
# $bit = &add_block_following
# $bit = &add_inline_following
# $bit = &uniq_inline_attributes
# $bit = &add_inline_within
# $bit = &copyfile
# $bit = &lint
# $bit = &get_combo_errors
# $bit = &get_xml_err_lines
# $bit = &load_lint_line_nums
# $bit = &get_pic_size
# $bit = &lose_idm_ids
# $bit = &lose_unicode_ents
# $bit = &norm_wd
# $bit = &hex2letter
# $bit = &utf82nimpho
# $bit = &get_ig_phons
# $bit = &wdcmp
# $bit = &wdsimplify
# $bit = &get_xml_header
# $bit = &get_xsl_header($projectname)
# $bit = &get_xsl_footer
# $hdr = &get_html_header("page title", "cssname");
# $sk = &get_idiom_sortkey($bit);
#
$SoundBase = "/usrdata3/audio";
$SOUND_DBASE = "/usrdata3/audio/DBASE.dat";

sub reduce
{
    my($e, $type) = @_;
    my($res, $eid);	
    if ($type =~ m|all|i)
    {
	$e = &del_xmlns($e);
	$e = &lose_idmids($e);
    }
    elsif ($type =~ m|xmlns|i)
    {
	$e = &del_xmlns($e);
    }
    elsif ($type =~ m|eids|i)
    {
	$e = &reduce_idmids($e);
    }
    return $e;
}

sub del_xmlns
{
    my($e) = @_;
    my($res, $eid);	
    $e =~ s| xmlns[^ =]*=\"[^\"]*\"||gi;
    return $e;
}

sub tidy_word
{
    my($wd) = @_;
    my($res, $eid);
    $wd =~ s| *(</st>) *|\1 |gi;
    $wd = restructure::lose_tag($wd, "st"); # lose the tags but not the contents	
    $wd =~ s|&\#x02C8;||go;
    $wd =~ s|&\#x02CC;||go;
    $wd =~ s|&\#x00B7;||go;
    $wd =~ s|&\#x2027;||go;
    #	    $wd =~ s|\'||go;
    $wd =~ s|&\#x2122;||go;
    $wd =~ s|&\#x2009;||go;
    $wd =~ s|&\#x00AE;||go;
    return $wd;
}

sub define_wdlist_links{
    $WDLISTLINK{"CEFR_A1"} = '<wordlist href="CEFR_A1" e:targetid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7ad7" e:targeteltid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7ad7" e:targetproject="WORDLISTS"/>';
    $WDLISTLINK{"CEFR_A2"} = '<wordlist href="CEFR_A2" e:targetid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7ad3" e:targeteltid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7ad3" e:targetproject="WORDLISTS"/>';
    $WDLISTLINK{"CEFR_B1"} = '<wordlist href="CEFR_B1" e:targetid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7acf" e:targeteltid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7acf" e:targetproject="WORDLISTS"/>';
    $WDLISTLINK{"CEFR_B2"} = '<wordlist href="CEFR_B2" e:targetid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7acb" e:targeteltid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7acb" e:targetproject="WORDLISTS"/>';
    $WDLISTLINK{"CEFR_C1"} = '<wordlist href="CEFR_C1" e:targetid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7ac7" e:targeteltid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7ac7" e:targetproject="WORDLISTS"/>';
    $WDLISTLINK{"CEFR_C2"} = '<wordlist href="CEFR_C2" e:targetid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7ac3" e:targeteltid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7ac3" e:targetproject="WORDLISTS"/>';
    $WDLISTLINK{"ox3000_new"} = '<wordlist href="ox3000_new" e:targetid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7adf" e:targeteltid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7adf" e:targetproject="WORDLISTS"/>';
    $WDLISTLINK{"OPAL_Spoken"} = '<wordlist href="OPAL_Spoken" e:targetid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7aee" e:targeteltid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7aee" e:targetproject="WORDLISTS"/>';
    $WDLISTLINK{"OPAL_spoken_phrases"} = '<wordlist href="OPAL_spoken_phrases" e:targetid="u6309bee381efbc42.-5dc40f8d.164d0ecc5fa.-189e" e:targeteltid="u6309bee381efbc42.-5dc40f8d.164d0ecc5fa.-189e" e:targetproject="WORDLISTS"/>';
    $WDLISTLINK{"OPAL_Written"} = '<wordlist href="OPAL_Written" e:targetid="ucc895cdac1ecc7f1.226ca1cb.164b1f48bde.7c8b" e:targeteltid="ucc895cdac1ecc7f1.226ca1cb.164b1f48bde.7c8b" e:targetproject="WORDLISTS"/>';
    $WDLISTLINK{"ox5000"} = '<wordlist href="ox5000" e:targetid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7adb" e:targeteltid="ucc895cdac1ecc7f1.226ca1cb.164b2f2c253.-7adb" e:targetproject="WORDLISTS"/>';
    $WDLISTLINK{"oxford_phrase_list"} = '<wordlist href="oxford_phrase_list" e:targetid="u14924562cdb56e04.1aa95c38.166e7ed0c4b.68e2" e:targeteltid="u14924562cdb56e04.1aa95c38.166e7ed0c4b.68e2" e:targetproject="WORDLISTS"/>';
}

sub get_indd_hdr
{
    my($res, $eid);
    $res = '<?xml version="1.0" encoding="UTF-8"?>&nl;<story xmlns:aid="http://ns.adobe.com/AdobeInDesign/4.0/" xmlns:aid5="http://ns.adobe.com/AdobeInDesign/5.0/">&nl;<dps-data>';
    $res =~ s|&nl;|\n|gi;
    return $res;
}

sub delete_unlinked_prongs
{
    my($e) = @_;
    my($res, $eid);	
    my($bit);
    my(@BITS);
    $e = restructure::rename_tag_in_tag($e, "vp-gs", "pron-gs", "vp-pron-gs");
    $e =~ s|(<pron-gs[ >].*?</pron-gs>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $target = restructure::get_tag_attval($bit, "pron-gs", "e:targeteltid"); 
	    if ($target =~ m|^ *$|)
	    {
		$bit = "";
	    }
	}
	$res .= $bit;
    }    
    $res = restructure::tag_rename($res, "vp-pron-gs", "pron-gs"); 
    return $res;
}

sub mycomm
{
    my($comm, $DBG, $quiet) = @_;
    unless ($quiet)
    {
	printf("%s\n", $comm);
    }
    unless ($DBG)
    {
	system($comm);
    }
    return 1;
}

sub clean_wd
{
    my($wd) = @_;
    $wd =~ s|&\#x02C8;||go;
    $wd =~ s|&\#x02CC;||go;
    $wd =~ s|&\#x00B7;||go;
    $wd =~ s|&\#x2027;||go;
    $wd =~ s|&\#x2122;||go;
    $wd =~ s|&\#x2009;||go;
    $wd =~ s|&\#x00AE;||go;
    return $wd;
}

sub get_tag_contents
{
    my($e, $tag) = @_;
    my($bit);
    if ($e =~ m|(<$tag[ >].*?</$tag>)|i)
    {
	$bit = $1;
	$bit =~ s|^ *<[^>]*>||i;
	$bit =~ s|</[^>]*> *$||i;
	return $bit;
    }
    return "";
}


sub get_tag_attval
{
    my($e, $tag, $att) = @_;
    my($bit);
    if ($e =~ m|(<$tag [^>]*>)|i)
    {
	$bit = $1;
	if ($bit =~ m| $att=\"(.*?)\"|i)
	{
	    return $1;
	}
    }
    return "";
}

sub get_attval
{
    my($e, $att) = @_;
    if ($e =~ m| $att=\"(.*?)\"|i)
    {
	return $1;
    }
    return "";
}

sub open_debug_files
{
    my($logfile);
    if ($LOG){
	if ($opt_L)
	{
	    $logfile = $opt_L;
	}
	else
	{
	    $logfile = sprintf("%s.log", $0);
	}
	open(log_fp, ">$logfile") || die "Unable to open >$logfile"; 
	binmode(log_fp, ":utf8");
    }
    if ($opt_I){
	open(bugin_fp, ">$0.in") || die "Unable to open >$0.in"; 
	binmode(bugin_fp, ":utf8");
    }    
    if ($opt_O){
	open(bugout_fp, ">$0.out") || die "Unable to open >$0.out"; 
	binmode(bugout_fp, ":utf8");
    }
}


sub open_debug_files_bug
{
    my($logfile);
#    open my $fh, ">:utf8", $filename or die "could not open $filename: $!\n";
#    open my $fh, "<:encoding(utf-8)", $filename or die "could not open $filename: $!\n";    
    #    if ($UTF8){
#    binmode(STDIN, ":utf8");
#    binmode(STDOUT, ":utf8");
    #    }
    if ($LOG){
	if ($opt_L)
	{
	    $logfile = $opt_L;
	}
	else
	{
	    $logfile = sprintf("%s.log", $0);
	}
	open log_fp, ">:utf8", ">$logfile" || die "Unable to open >$logfile"; 
    }
    if ($opt_I){
	    open bugin_fp,  ">:utf8", ">$0.in" || die "Unable to open >$0.in"; 
    }    
    if ($opt_O){
	    open bugout_fp,  ">:utf8", ">$0.out" || die "Unable to open >$0.out"; 
    }
}

sub close_debug_files
{
    if ($LOG){
	close(log_fp);
    }
    if ($opt_I){
	close(bugin_fp);
    }
    if ($opt_O){
	close(bugout_fp);
    }
}

sub get_phon_info
{
    my($e) = @_;
    my ($wd, $file, $phon, $lang);

    $wd = "";
    $file = "";
    $lang = "";
    $phon = "";
    if ($e =~ m| wd=\"(.*?)\"|io)
    {
	$wd = $1;
    }
    if ($e =~ m| file=\"(.*?)\"|io)
    {
	$file = $1;
    }
    $phon = $e;
    $phon =~ s|</?phon.*?>||gio;
    $phon =~ s|<\?xm-replace_text.*?\?>||gio;
    if ($e =~ m|<phon-gb|)
    {
	$lang = "gb";
    }
    else
    {
	$lang = "us";
    }
    return ($file, $lang, $phon, $wd);
}

sub get_name_date
{
    my ($date, $name, $comment);
    $date = scalar localtime;
    $name =  getlogin || (getpwuid($<)) [0];
    $comment = sprintf("<!-- Created by: %s Date: %s -->", $name, $date);
    return $comment;
}

sub my_name_date
{
    my ($date, $name, $comment);
    $date = scalar localtime;
    $name =  getlogin || (getpwuid($<)) [0];
    $comment = sprintf("Created by: %s Date: %s", $name, $date);
    return $comment;
}

sub get_pwd
{
    my($pwd);
    $pwd = abs_path($0);
    $pwd =~ s|/[^/]*$||;
    return $pwd;
}

sub gmt
{
    my($gmt) = @_;
    my($gmt_s, $gmt_m, $gmt_h, $gmt_md, $gmt_mon, $gmt_y, $gmt_d, $gmt_yd, $gmt_j);
    if ($gmt =~ m|^ *$|)
    {
	$gmt = time;
    }
    ($gmt_s, $gmt_m, $gmt_h, $gmt_md, $gmt_mon, $gmt_y, $gmt_d, $gmt_yd, $gmt_j) = gmtime($gmt);
    $gmt = sprintf("%02d:%02d:%02d\n", $gmt_h, $gmt_m, $gmt_s); 
    return $gmt;
}
##
sub get_xyz_file
{
    my($file, $BASE) = @_;
    my($fname, $fname2, $fcp, $resdir, $resfile);
    my($d1, $d2, $d3);
    $fcp = $file;
    $file =~ s|^.*/||;
    $file =~ tr|A-Z|a-z|;
    $file =~ s|^aux|xaux|;
    $file =~ s|^con|xcon|;
    $file =~ s|^nul|xnul|;
    $file =~ s| +& +|and|gio;
    $df = $file;
    $df =~ s|_.*$||;
    $df =~ s|$|__|;
    $df =~ s| +|_|gio;
    $df =~ s|[^a-z\_]+||gio;
    $file =~ s|[^a-z0-9\.\_]+||gio;
    if ($df =~ m|^(((.).).)|)
    {
	$d3 = $1;
	$d2 = $2;
	$d1 = $3;
	$resdir = sprintf("%s/%s/%s/%s", $BASE, $d1, $d2, $d3);
	$resfile = sprintf("%s/%s", $resdir, $file)
    }
    $resdir =~ tr|A-Z|a-z|;
    $resfile =~ tr|A-Z|a-z|;
    return ($resdir, $resfile);
}
##


sub get_fileh
{
    my($e) = @_;
    my($bit, $h, @BITS);
    $e =~ s|<\!--.*?-->||goi;
    $e =~ tr|A-Z|a-z|;
    $e =~ s|</?hs>||goi;
    $e =~ s|<ht>|<h>|goi;
    $e =~ s|</ht>|</h>|goi;
    $e =~ s|<hti>.*?</hti> *| |goi;
    $e =~ s|<hr>.*?</hr> *| |goi;
    # lose funnies
    $e =~ s|&[swp];||gio;
    $e =~ s|\$||gio;
    $e =~ s|&middot;||gio;
    $e =~ s|&trade;||gio;
    $e =~ s|&thinsp;||gio;
    $e =~ s|&[lr][sd]quo;| |gio;
    $e =~ s|<h> *|<h>|goi;
    if ($e =~ /<h> *(.*?)<\/h>/i)
    {
	$h = $1;
	if ($h =~ /&/)
	{
	    $h = &change_ae_ents($h);
	    $h = &ents2letter($h);
	}
	if ($h =~ /^ *$/)
	{
	    $h = "dk";
	}
	$h =~ s|^[\$\-]*||o;
	$h =~ s|^the ||o;
	$h =~ s|[^a-z0-9].*||io;
	return $h;
    }
}

sub get_target_fname
{
    my($h) = @_;
    my($fname);
    my($comm);
    my($firstletter);
    my($dir);
    $fname = $h;
    $fname =~ s|<hr>.*?</hr>||gio;
    $fname =~ s|<hti>.*?</hti>||gio;
    $fname =~ s|^[^a-z0-9]*||gio;
    $fname =~ tr|A-Z|a-z|;
    $fname =~ s|[^a-z0-9].*||;
    #    $fname =~ s|\-.*||o;
    $firstletter = $fname;
    $firstletter =~ s|^(.).*$|$1|o;
    
    $dir = sprintf("%s/%s", $dictfiles, $firstletter);
    
    unless (-d $dir)
    {
	$comm = sprintf("mkdir -p %s", $dir);
	system($comm);
    }	
    $fname = sprintf("%s/%s.sgm", $dir, $fname);	
    $lockfile = $fname;
    $lockfile =~ s|/dictfiles/|/lockfiles/|o;
    $lockfile =~ s|\.sgm *$|\.lck|o;
    return($fname, $lockfile);
}

sub change_ae_ents
{
    my ($line) = @_;
    $line =~ s|&\#128;|&Auml;|go;
    $line =~ s|&\#129;|&Aring;|go;
    $line =~ s|&\#130;|&Ccedil;|go;
    $line =~ s|&\#131;|&Eacute;|go;
    $line =~ s|&\#132;|&Ntilde;|go;
    $line =~ s|&\#133;|&Ouml;|go;
    $line =~ s|&\#134;|&Uuml;|go;
    $line =~ s|&\#135;|&aacute;|go;
    $line =~ s|&\#136;|&agrave;|go;
    $line =~ s|&\#137;|&acirc;|go;
    $line =~ s|&\#138;|&auml;|go;
    $line =~ s|&\#139;|&atilde;|go;
    $line =~ s|&\#140;|&aring;|go;
    $line =~ s|&\#141;|&ccedil;|go;
    $line =~ s|&\#142;|&eacute;|go;
    $line =~ s|&\#143;|&egrave;|go;
    $line =~ s|&\#144;|&ecirc;|go;
    $line =~ s|&\#145;|&euml;|go;
    $line =~ s|&\#146;|&iacute;|go;
    $line =~ s|&\#147;|&igrave;|go;
    $line =~ s|&\#148;|&icirc;|go;
    $line =~ s|&\#149;|&iuml;|go;
    $line =~ s|&\#150;|&ntilde;|go;
    $line =~ s|&\#151;|&oacute;|go;
    $line =~ s|&\#152;|&ograve;|go;
    $line =~ s|&\#153;|&ocirc;|go;
    $line =~ s|&\#154;|&ouml;|go;
    $line =~ s|&\#155;|&otilde;|go;
    $line =~ s|&\#156;|&uacute;|go;
    $line =~ s|&\#157;|&ugrave;|go;
    $line =~ s|&\#158;|&ucirc;|go;
    $line =~ s|&\#159;|&uuml;|go;
    $line =~ s|&\#163;|&pound;|go;
    $line =~ s|&\#170;|&ordf;|go;
    $line =~ s|&\#174;|&AElig;|go;
    $line =~ s|&\#190;|&aelig;|go;
    $line =~ s|&\#191;|&oslash;|go;
    $line =~ s|&\#192;|&iquest;|go;
    $line =~ s|&\#193;|&iexclam;|go;
    $line =~ s|&\#201;|&hellip;|go;
    $line =~ s|&\#203;|&Agrave;|go;
    $line =~ s|&\#204;|&Atilde;|go;
    $line =~ s|&\#205;|&Otilde;|go;
    $line =~ s|&\#210;|&ldquo;|go;
    $line =~ s|&\#211;|&rdquo;|go;
    $line =~ s|&\#212;|&lsquo;|go;
    $line =~ s|&\#213;|&rsquo;|go;
    $line =~ s|&\#216;|&yuml;|go;
    $line =~ s|&\#229;|&Acirc;|go;
    $line =~ s|&\#230;|&Ecirc;|go;
    $line =~ s|&\#231;|&Aacute;|go;
    $line =~ s|&\#232;|&Euml;|go;
    $line =~ s|&\#233;|&Egrave;|go;
    $line =~ s|&\#234;|&Iacute;|go;
    $line =~ s|&\#235;|&Icirc;|go;
    $line =~ s|&\#236;|&Iuml;|go;
    $line =~ s|&\#237;|&Igrave;|go;
    $line =~ s|&\#238;|&Oacute;|go;
    $line =~ s|&\#239;|&Ocirc;|go;
    $line =~ s|&\#241;|&Ograve;|go;
    $line =~ s|&\#242;|&Uacute;|go;
    $line =~ s|&\#243;|&Ucirc;|go;
    $line =~ s|&\#244;|&Ugrave;|go;
    $line =~ s|Ž|&eacute;|go;
    $line =~ s|‡|&aacute;|go;
    $line =~ s|’|&iacute;|go;
    $line =~ s|À|&iquest;|go;
    $line =~ s|–|&ntilde;|go;
    $line =~ s|—|&oacute;|go;
    $line =~ s|Á|&iexclam;|go;
    $line =~ s|œ|&uacute;|go;
    $line =~ s|É|&hellip;|go;
    $line =~ s|Ÿ|&uuml;|go;
    $line =~ s|ƒ|&Eacute;|go;
    return($line);
}

sub ents2letter
{
    my($e) = @_;
    $e =~ s|&\#x0026;|and|gi;
    $e = &lose_unicode_ents($e);
    if ($e =~ m||)
    {
	$e = &hex2letter($e);
    }
    $e =~ s|&[swp];||g;
    $e =~ s|&[lr][sd]quo;| |gio;
    $e =~ s|&nbhyph;|-|gio;
    $e =~ s| *&hellip; *| |gio;
    $e =~ s|&iexclam;| |go;
    $e =~ s|&iquest;| |go;
    $e =~ s| *&amp; *| and |gio;
    $e =~ s|&trade;||gio;
    $e =~ s|&tm;||gio;
    $e =~ s|&copy;||gio;
    $e =~ s|&reg;||gio;
    $e =~ s|&szlig;|ss|gio;
    $e =~ s| *&[^&; ]*thinsp[^&; ]*;||gio;
    $e =~ s| *&[^&; ]*hdwdsp[^&; ]*;||gio;
    $e =~ s|^ *||go;
    $e =~ s|&(.).*?;|$1|go;
    return $e;
}

sub get_dr_old
{
    my($e) = @_;
    my($dr);
    $e = &lose_unicode_ents($e);
    $e =~ s|</ht>|</h>|gio;
    $e =~ s|</?audio[^>]*>| |gio;
    $e =~ s|&[spw]b?;||go;
    $e =~ s|&middot;||go;
    $e =~ s|&trade;||go;
    $e =~ s|&thinsp;||go;
    $e =~ s|&[lr][sd]quo;||goi;
    $e =~ s|&[spw];||go;
    $e =~ s|<(/?)zd([ >])|<$1dr$2|gi;
    if ($e =~ m|(<dr[ >].*?</dr *>)|i)
    {
	$dr = $1;
	$dr =~ s|<.*?>||gio;
	$dr = &change_ae_ents($dr);
	$dr = &ents2letter($dr);
	return $h;
    }
}

sub get_hex_h
{
    my($e, $format, $clean) = @_;
    my($h);
    $e =~ s|<ht([> ])|<h$1|gio;
    $e =~ s|</ht>|</h>|gio;
    if ($e =~ m|(<h[ >].*?</h *>)|i)
    {
	$h = $1;
	$h =~ s|&#x0026;|and|gi;
	unless ($format =~ m|hex|i)
	{
	    $h = &hex2letter($h);
	    $h = &lose_unicode_ents($h);
	}
	$h =~ s|<.*?>||gio;
	#	$h =~ s|,.*||;
	$h =~ s|</?audio[^>]*>| |gio;
	if ($clean)
	{
	    $h =~ s|&\#x02C8;||go;
	    $h =~ s|&\#x02CC;||go;
	    $h =~ s|&\#x00B7;||go;
	    $h =~ s|&\#x2027;||go;
	    #	    $h =~ s|\'||go;
	    $h =~ s|&\#x2122;||go;
	    $h =~ s|&\#x2009;||go;
	    $h =~ s|&\#x00AE;||go;
	}
	if (0)
	{
	    $h =~ s|&\#x02C8;||go;
	    $h =~ s|&\#x02CC;||go;
	    $h =~ s|&\#x00B7;||go;
	    $h =~ s|\'||go;
	    $h =~ s|&\#x2122;||go;
	    $h =~ s|&\#x2009;||go;
	    $h =~ s|&\#x00AE;||go;
	    ## quotes
	    $h =~ s|&\#x00AB;||go;
	    $h =~ s|&\#x201C;||go;
	    $h =~ s|&\#x201E;||go;
	    $h =~ s|&\#x2018;||go;
	    $h =~ s|&\#x201A;||go;
	    $h =~ s|&\#x0022;||go;
	    $h =~ s|&\#x00BB;||go;
	    $h =~ s|&\#x201D;||go;
	    $h =~ s|&\#x201C;||go;
	    $h =~ s|&\#x2019;||go;
	    $h =~ s|&\#x2018;||go;
	    $h =~ s|&[sp]stress;||gio;
	    $h =~ s|<stress.*?</stress[^>]*>||gi;
	}
	return $h;
    }
}

sub destress
{
    my($e) = @_;
    $e =~ s|&\#x02C8;||gio;
    $e =~ s|&\#x02CC;||gio;
    $e =~ s|&\#x00B7;||gio;
    $e =~ s|&\#x2027;||gio;
    $e =~ s|&\#x2122;||gi;
    $e =~ s|&\#x2019;|\'|gi;
    return $e;
}

sub get_dps_entry_id
{
    my($e) = @_;
    my ($res);
    $res = "";
    if (m|<entry[^>]* e:id=\"(.*?)\"|i)
    {
	$res = $1;
    }
    return $res;
}

sub get_entry_id
{
    my($e) = @_;
    my($res);
    $res = "";
    if (m|<entry[^>]* e?id=\"(.*?)\"|i)
    {
	$res = $1;
    }
    return $res;
}


sub get_h
{
    my($e) = @_;
    my($h);
    $e = &lose_unicode_ents($e);
    $e =~ s|<ht([> ])|<h$1|gio;
    $e =~ s|</ht>|</h>|gio;
    $e =~ s|</?audio[^>]*>| |gio;
    $e =~ s|&[spw]b?;||go;
    $e =~ s|&middot;||go;
    $e =~ s|\'||go;
    $e =~ s|&trade;||go;
    $e =~ s|&[^&;]*thinsp;||go;
    $e =~ s|&[lr][sd]quo;||goi;
    $e =~ s|&[spw];||go;
    $e =~ s|&[sp]stress;||gio;
    if ($e =~ m|(<h[ >].*?</h *>)|i)
    {
	$h = $1;
	$h =~ s|<.*?>||gio;
	$h =~ s|,.*||;
	$h = &change_ae_ents($h);
	$h = &ents2letter($h);
	return $h;
    }
}

sub get_dr
{
    my($e, $zd) = @_;
    my($h);
    $e = &lose_unicode_ents($e);
    $e =~ s|</?audio[^>]*>| |gio;
    $e =~ s|&[spw]b?;||go;
    $e =~ s|&middot;||go;
    $e =~ s|\'||go;
    $e =~ s|&trade;||go;
    $e =~ s|&[^&;]*thinsp;||go;
    $e =~ s|&[lr][sd]quo;||goi;
    $e =~ s|&[spw];||go;
    $e =~ s|&[sp]stress;||gio;
    if ($zd)
    {
	$e =~ s|<(/?)zd([ >])|<$1dr$2|gi;
    }
    if ($e =~ m|(<dr[ >].*?</dr *>)|i)
    {
	$h = $1;
	$h =~ s|<.*?>||gio;
	$h =~ s|,.*||;
	$h = &change_ae_ents($h);
	$h = &ents2letter($h);
	return $h;
    }
}

sub get_xml_h
{
    my($e, $az) = @_;
    my($h);
    unless ($az =~ m|^ *$|)
    {
	$e = &lose_unicode_ents($e);
    }
    $e =~ s|<hm[ >].*?</hm>||gio;
    #    $e =~ s|<.*?>||gio;
    $e =~ s| *$||;
    $e =~ s|<ht([> ])|<h$1|gio;
    $e =~ s|</ht>|</h>|gio;
    $e =~ s|</?audio[^>]*>| |gio;
    $e =~ s|&[spw]b?;||go;
    $e =~ s|&middot;||go;
    $e =~ s|\'||go;
    $e =~ s|&trade;||go;
    $e =~ s|&[^&;]*thinsp;||go;
    $e =~ s|&[lr][sd]quo;||goi;
    $e =~ s|&[spw];||go;
    $e =~ s|&[sp]stress;||gio;
    if ($e =~ m|(<h[ >].*?</h *>)|i)
    {
	$h = $1;
	$h =~ s|<.*?>||gio;
	$h =~ s|,.*||;
	unless ($az =~ m|^ *$|)
	{
	    $h = &change_ae_ents($h);
	    $h = &ents2letter($h);
	}
	$h = &tidy_hexwd($h);
	return $h;
    }
}

sub get_norm_hdwd
{
    my($e) = @_;
    my($h);
    if ($e =~ m|(<h[ >].*?</h>)|i)
    {
	$h = $1;
	$h =~ s|<.*?>||gi;
	$h =~ s|&[swp];||gi;
	$h =~ s|&\#x02C8;||go;
	$h =~ s|&\#x02CC;||go;
	$h =~ s|&\#x00B7;||go;
	$h = &norm_wd($h);
	$h =~ tr|A-Z|a-z|;
	return $h;
    }
    return "";
}

sub get_hdwd
{
    my($e) = @_;
    my($h);
    $e =~ s|<ht([ >])|<h$1|gio;
    $e =~ s|</ht>|</h>|gio;
    if ($e =~ m|(<h[ >].*?</h>)|i)
    {
	$h = $1;
	if ($h =~ m|&\#x|i)
	{
	    $h = &hex2letter($h);
	    $h = &lose_unicode_ents($h);
	}
	$h =~ s|&[spw]b?;||go;
	$h =~ s|&iexcl;||go;
	$h =~ s|\!||go;
	$h =~ s|&middot;||go;
	$h =~ s|&trade;||go;
	$h =~ s|&thinsp;||go;
	$h =~ s|&[lr]squo;|\'|goi;
	$h =~ s|&[lr][sd]quo;||goi;
	$h =~ s|<.*?>||gio;
	$h =~ s|,.*$||;
	$h = &change_ae_ents($h);
	$h = &ents2letter($h);
	return $h;
    }
    else
    {
	return $file_hdwd;
    }
}


sub get_hdwd_BAK
{
    my($e) = @_;
    my($h);
    $e = &hex2letter($e);
    $e = &lose_unicode_ents($e);
    $e =~ s|<ht>|<h>|gio;
    $e =~ s|</ht>|</h>|gio;
    $e =~ s|</?audio[^>]*>||gio;
    $e =~ s|&[spw]b?;||go;
    $e =~ s|&middot;||go;
    $e =~ s|&trade;||go;
    $e =~ s|&thinsp;||go;
    $e =~ s|&[lr][sd]quo;||goi;
    $e =~ s|&[spw];||go;
    if ($e =~ m|(<h[ >].*?</h>)|i)
    {
	$h = $1;
	$h =~ s|<.*?>||gio;
	$h =~ s|,.*$||;
	$h = &change_ae_ents($h);
	$h = &ents2letter($h);
	return $h;
    }
    else
    {
	return $file_hdwd;
    }
}


sub get_accented_hdwd
{
    my($e) = @_;
    my($h);
    #    $e = &hex2letter($e);
    $e = &lose_unicode_ents($e);
    $e =~ s|<ht>|<h>|gio;
    $e =~ s|</ht>|</h>|gio;
    $e =~ s|</?audio[^>]*>||gio;
    $e =~ s|&[spw]b?;||go;
    $e =~ s|&middot;||go;
    $e =~ s|&trade;||go;
    $e =~ s|&thinsp;||go;
    $e =~ s|&[lr][sd]quo;||goi;
    $e =~ s|&[spw];||go;
    if ($e =~ m|(<h[ >].*?</h>)|i)
    {
	$h = $1;
	$h =~ s|<.*?>||gio;
	$h =~ s|,.*$||;
	$h = &change_ae_ents($h);
	#	$h = &ents2letter($h);
	return $h;
    }
    else
    {
	return $file_hdwd;
    }
}


sub get_dr_old2
{
    my($e, $zd) = @_;
    my($h);
    $e = &hex2letter($e);
    $e = &lose_unicode_ents($e);
    $e =~ s|</?audio[^>]*>||gio;
    $e =~ s|&[spw]b?;||go;
    $e =~ s|&middot;||go;
    $e =~ s|&trade;||go;
    $e =~ s|&thinsp;||go;
    $e =~ s|&[lr][sd]quo;||goi;
    $e =~ s|&[spw];||go;
    if ($e =~ m|(<dr[ >].*?</dr>)|i)
    {
	$h = $1;
	$h =~ s|<.*?>||gio;
	$h = &change_ae_ents($h);
	$h = &ents2letter($h);
	return $h;
    }
    else
    {
	return "";
    }
}

sub get_h_hm
{
    my($e) = @_;
    my($res);
    my($h, $hm);
    $res = "";
    $h = "";
    $hm = "";
    $e = &lose_unicode_ents($e);
    $e =~ s|\t| |gio;
    $e =~ s|HT>|H>|gio;
    $e =~ s|HTI>|HR>|gio;
    if ($e =~ m|(<H[> ].*?</H>)|i)
    {
	$h = $1;
	$h =~ s|<HR>.*?</HR>||gio;
	$h =~ s|<[^>]*>||gio;
	$h =~ s|^ *||gio;
    }
    if ($e =~ m|<HM>(.*?)</HM>|i)
    {
	$hm = $1;
	$hm =~ s|^ *||gio;
    }
    $res = sprintf("<h>%s</h><hm>%s</hm>", $h, $hm);
    return $res;
}

sub get_topg
{
    my($e) = @_;
    my($res);
    if ($e =~ m|(<top-g.*?</top-g>)|io)
    {
	$res = $1;
    }
    return $res;
}

sub get_posg
{
    my($e) = @_;
    my($pres, $pbit, $ptag);
    my @PBITS;
    undef %PTMP;
    $e =~ s|(<p .*?>)|&split;&fk;$1&split;|gio;
    @PBITS = split(/&split;/, $e);
    $pres = "";
    foreach $pbit (@PBITS)
    {
	if ($pbit =~ s|&fk;||gio)
	{
	    if ($pbit =~ m| p=\"(.*?)\"|io)
	    {
		$ptag = $1;
		$PTMP{$ptag} = 1;
	    }
	}
    }
    foreach $ptag (sort keys %PTMP)
    {
	$pres = sprintf("%s%s\t", $pres, $ptag);
    }
    $pres =~ s|\t *$||;
    return $pres;
}

sub get_xh_xhm
{
    my($e) = @_;
    my($res, $xh, $xhm, @XRS, $xr, $h_hm);
    $res = "";
    $xh = "";
    $xhm = "";
    $e =~ s|\t| |gio;
    return unless ($e =~ s|<XR |&split;<XR |gio);
    @XRS = split(/&split;/, $e);
  xrfloop:
    foreach $xr (@XRS)
  {
      next xrfloop unless ($xr =~ /<XR/i);
      if ($xr =~ m|<XHI?>(.*?)</XHI?>|i)
      {
	  $xh = $1;
	  $xh =~ s|^ *||gio;
      }
      if ($xr =~ m|<XHM>(.*?)</XHM>|i)
      {
	  $xhm = $1;
	  $xhm =~ s|^ *||gio;
      }
      $res = sprintf("%s<H>%s</H><HM>%s</HM>", $res, $xh, $xhm);
  }
    return $res;
}


sub get_sortkey
{
    my($e) = @_;
    my($h, $h2, $sk);
    my($hm);
    $hm = "";
    $h = "";
    $h = restructure::get_tag_contents($e, "h"); 
    if ($e =~ m|<hm[ >]|)
    {
	$hm = restructure::get_tag_contents($e, "hm"); 
    } 
    else {
	$hm = restructure::get_tag_attval($e, "entry", "hm"); 
    }
    $h = &lose_unicode_ents($h);
    $h =~ s|&[spw]b?;||go;
    $h =~ s|&middot;||go;
    $h =~ s|&[spw];||go;
    $h =~ s|&(.).*?;|$1|go;
    $h =~ s|^the ||oi;
    $h =~ s|\'||oi; # ignore apostrophe in sort key ...
    $h = restructure::tag_delete($h, "st"); # the The tag at the start
    $h2 = $h;
    if ($h2 =~ s|^\-+||o)
    {
	$h2 = sprintf("%s-", $h2);
    }
    $h =~ s|^\-*||o;
    $h =~ s|\-||go;
    $h =~ s|,.*||o;
    $h =~ s|[^a-zA-Z]||go;       
    $h =~ tr|A-Z|a-z|;
    #    $psk = sprintf("%s#%s#%s", $h, $h2, $hm);
    $sk = sprintf("%-25s%-5s%-25s", $h, $h2, $hm);
    return $sk;
}

sub get_sortwd
{
    my($e) = @_;
    my($h, $h2, $sk);
    my($hm);
    $hm = "";
    $h = $e;
    
    $h = &lose_unicode_ents($h);
    $h =~ s|&[spw]b?;||go;
    $h =~ s|&middot;||go;
    $h =~ s|&[spw];||go;
    $h =~ s|&(.).*?;|$1|go;
    $h =~ s|^the ||oi;
    $h =~ s|\'||oi; # ignore apostrophe in sort key ...
    $h2 = $h;
    if ($h2 =~ s|^\-+||o)
    {
	$h2 = sprintf("%s-", $h2);
    }
    $h =~ s|^\-*||o;
    $h =~ s|\-||go;
    $h =~ s|,.*||o;
    $h =~ s|[^a-zA-Z]||go;       
    $h =~ tr|A-Z|a-z|;
    #    $psk = sprintf("%s#%s#%s", $h, $h2, $hm);
    $sk = sprintf("%-25s%-5s", $h, $h2);
    return $sk;
}


sub def_prevletter
{
    my($letter) = @_;
    $prevletter{"a"} = "\'";
    $prevletter{"b"} = "a";
    $prevletter{"c"} = "b";
    $prevletter{"d"} = "c";
    $prevletter{"e"} = "d";
    $prevletter{"f"} = "e";
    $prevletter{"g"} = "f";
    $prevletter{"h"} = "g";
    $prevletter{"i"} = "h";
    $prevletter{"j"} = "i";
    $prevletter{"k"} = "j";
    $prevletter{"l"} = "k";
    $prevletter{"m"} = "l";
    $prevletter{"n"} = "m";
    $prevletter{"o"} = "n";
    $prevletter{"p"} = "o";
    $prevletter{"q"} = "p";
    $prevletter{"r"} = "q";
    $prevletter{"s"} = "r";
    $prevletter{"t"} = "s";
    $prevletter{"u"} = "t";
    $prevletter{"v"} = "u";
    $prevletter{"w"} = "v";
    $prevletter{"x"} = "w";
    $prevletter{"y"} = "x";
    $prevletter{"z"} = "y";
}

sub def_nextletter
{
    my($letter) = @_;
    $nextletter{"a"} = "b";
    $nextletter{"b"} = "c";
    $nextletter{"c"} = "d";
    $nextletter{"d"} = "e";
    $nextletter{"e"} = "f";
    $nextletter{"f"} = "g";
    $nextletter{"g"} = "h";
    $nextletter{"h"} = "i";
    $nextletter{"i"} = "j";
    $nextletter{"j"} = "k";
    $nextletter{"k"} = "l";
    $nextletter{"l"} = "m";
    $nextletter{"m"} = "n";
    $nextletter{"n"} = "o";
    $nextletter{"o"} = "p";
    $nextletter{"p"} = "q";
    $nextletter{"q"} = "r";
    $nextletter{"r"} = "s";
    $nextletter{"s"} = "t";
    $nextletter{"t"} = "u";
    $nextletter{"u"} = "v";
    $nextletter{"v"} = "w";
    $nextletter{"w"} = "x";
    $nextletter{"x"} = "y";
    $nextletter{"y"} = "z";
    $nextletter{"z"} = "z";
}

sub get_entry_file
{
    my($fname) = @_;
    $entry_file = `cat $fname`;    
    $entry_file =~ s|\n| |go;	# 
    $entry_file =~ s|||go;
    $entry_file =~ s|(<[^>]*)$|$1 |o;
    $entry_file =~ s|</ENTRY> *|</ENTRY>\n|goi;
    $entry_file =~ s| *<ENTRY|<ENTRY|goi;
    $entry_file =~ s|\n\n|\n|go;
    $entry_file =~ s|&middot;|&w;|go;
    return $entry_file;
}

sub get_date
{
    my($date);
    $date = `date +%y%m%d`;
    chomp $date;
    return $date;
}

sub get_date2
{
    my($date);
    $date = `date +%d/%m/%y`;
    chomp $date;
    return $date;
}

sub get_date3
{
    my($date);
    $date = scalar localtime;
    return $date;
}

sub get_logfile
{
    my($f) = @_;
    my($ct);
    my($date) = &get_date;
    $basef = $f;
    $f = sprintf("%s_%s.log", $f, $date);

    
    while (-f $f)
    {
	$ct++;
	$f = sprintf("%s_%s_%s.log", $basef, $date, $ct);
    }
    return $f;
}
sub open_logfile
{
    my($f, $mode) = @_;
    my($comment);
    if ($mode =~ m|o|)
    {
	open(log_fp, ">$f") || die "Unable to open >$f"; 
    }
    else
    {
	open(log_fp, ">>$f") || die "Unable to open >>$f"; 
    }
    $comment = &get_name_date;
    printf(log_fp "%s\n", $comment); 
    printf(STDERR "Log file opened: %s\n", $f); 
}

sub get_sound_dir
{
    my($wd, $vanilla) = @_;
    my($letters, $d1, $d2, $d3, $subdir, $add_the);
    $wd =~ s|^con|xcon|;
    # may not work as wanted the next one - aim is that the word with become "the accused" -> accused so that they are in the same folder with or without "the" at the start of the word
    unless ($vanilla)
    {
	$wd =~ s|^the[ _]||;    
	$letters = sprintf("%s__", $wd);	    
	$letters =~ s|\#|__\#|;   
    }
    if ($letters =~ m|^(((.).).)|)
    {
	$d1 = $3;
	$d2 = $2;
	$d3 = $1;
	$subdir = sprintf("%s/%s/%s", $d1, $d2, $d3);
    }
    return $subdir;
}

sub mkdir
{
    my($dir, $log) = @_;
    my($comm);
    unless (-d $dir)
    {
	$comm = sprintf("mkdir -p %s", $dir);
	system($comm);
	if ($log)
	{
	    printf(log_fp "%s\n", $comm); 
	}
    }	
}

sub tidywd
{
    my($wd, $uscore) = @_;
    $wd =~ tr|A-Z|a-z|;
    # get the entities to the base letter
    $wd = &hex2letter($wd);
    $wd =~ s|<pnc[^>]*>.*?</pnc>||gi;
    $wd =~ s|,.*|$1|gio;
    $wd =~ s|&(.)acute;|$1|gio;
    $wd =~ s|&(.)grave;|$1|gio;
    $wd =~ s|&(.)cedil;|$1|gio;
    $wd =~ s|&(.)circ;|$1|gio;
    $wd =~ s|&(.)uml;|$1|gio;
    $wd =~ s|&(.)tilde;|$1|gio;
    $wd =~ s|&trade;||gio;
    $wd =~ s|&tm;||gio;
    $wd =~ s|&copy;||gio;
    $wd =~ s|&reg;||gio;
    $wd =~ s|&.*?;||gio;
    $wd =~ s|</?st.*?>||gi;
    $wd =~ s|</?ptl.*?>||gi;
    if ($uscore =~ m|_|)
    {
	$wd =~ s|[^a-z0-9_]+|_|g;
    }
    else
    {
	$wd =~ s|[^a-z0-9_]+||g;
    }
    return $wd;
}

sub tidy_hexwd
{
    my($e) = @_;
    $e =~ s|&\#x2122;||gio; # &tm;
    $e =~ s|&\#x2022;||gio; # &tm;
    $e =~ s|&\#x02c8;||gio;
    $e =~ s|&\#x02cc;||gio;
    $e =~ s|&\#x00b7;||gio;
    $e =~ s|  +| |g;
    $e =~ s|^ *||g;
    $e =~ s| *$||g;
    return $e;
}

sub add_space_between
{
    my($e, $a, $b) = @_;
    my($bit, $res);
    $e =~ s|(</$a *> *<$b)([ >])|$1 $2|gi;
    return $e;
}

sub add_block_following
{
    my($e, $a, $b) = @_;
    my($bit, $res);
    $e =~ s|(</$a *> *<$b)([ >])|$1 display=\"block\"$2|gi;
    return $e;
}

sub add_inline_following
{
    my($e, $a, $b) = @_;
    my($bit, $res);
    $e =~ s|(</$a *> *<$b)([ >])|$1 display=\"inline\"$2|gi;
    return $e;
}


sub uniq_inline_attributes
{
    my($e) = @_;
    my($bit, $res);
    my(@BITS);
    unless ($e =~ m|<[^>]* display=[^>]*display=|)
    {
	return $e;
    }
    $e =~ s|(<[^/][^>]+>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    while ($bit =~ s|( display=\".*?\".* )display=\"[^\"]*\"|$1|)
	    {}
	}
	$res .= $bit;
    }
    return $res;
}

sub add_inline_within
{
    my($e, $parent, $child) = @_;
    my($bit, $res);
    my(@BITS);
    unless ($e =~ m|<$parent|i)
    {
	return $e;
    }
    $e =~ s|(<$parent[ >].*?</$parent>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    $bit =~ s|<$child([ >])|<$child display=\"inline\"$1|gi;
	}
	$res .= $bit;
    }
    return $res;
}

sub copyfile
{
    my($f, $res) = @_;
    $comm = sprintf("cp \"%s\" \"%s\"", $f, $res);
    if (-d $res)
    {
	$resf = $f;
	$resf =~ s|^.*/||;
	$resf = sprintf("%s/%s", $res, $resf);
	if (compare($f, $resf))
	{
	    system($comm);
	}
    }
    else
    {
	system($comm);
    }
}

sub lint
{
    my($xml_f, $combo, $dbg) = @_; 
    my($xmllint, $comm, $xml_errs_f, $lint_f, $combo_errs_f);
    $xmllint = "/usr/local/bin/xmllint --noout";
    $lint_f = sprintf("%s.errs", $xml_f);
    $comm  = sprintf("$xmllint $xml_f 2> $lint_f");
    $pcomm  = sprintf("$xmllint $xml_f > $lint_f");
    printf("\n$pcomm\n\n"); 
    unless ($dbg)
    {
	system($comm);
	if (-s $lint_f)
	{
	    $xml_errs_f = sprintf("%s.errs", $xml_f);
	    $combo_errs_f = sprintf("%s.errs", $combo);
	    printf(STDERR "\n\nWARNING: There are XML errors in %s [see %s, %s, %s for details]\n\n", $xml_f, $lint_f, $xml_errs_f, $combo_errs_f); 
	    &load_lint_line_nums($lint_f);
	    &get_xml_err_lines($xml_f, $xml_errs_f);
	    unless ($combo =~ m|^ *$|)
	    {
		#		printf("\n\n$combo ERRORS\n\n"); 
		
		&get_combo_errors($combo, $combo_errs_f);
	    }
	    return ($combo_errs_f);
	}
    }    
    return 0;
}

sub get_combo_errors
{
    my($f, $out_f) = @_;
    my $id;
    open(in_fp, "$f") || die "Unable to open $f"; 
    open(out_fp, ">$out_f") || die "Unable to open >$out_f"; 
    while (<in_fp>)
    {
	chomp;
	if (m|<entry[^>]*id=\"(.*?)\"|io)
	{
	    $id = $1;
	    if ($ERR_ID{$id})
	    {
		printf(out_fp "%s\n", $_);
	    }
	}
    }
    close(in_fp);
    close(out_fp);
} 


# print the error lines from XML and store the IDS that fail
sub get_xml_err_lines
{
    my($f, $out_f) = @_;
    my($lct, $id, %ERR_ID);
    open(lint_fp, "$f") || die "Unable to open $f"; 
    open(out_fp, ">$out_f") || die "Unable to open >$out_f"; 
    while (<lint_fp>)
    {
	chomp;
	if ($LINT_ERR{++$lct})
	{
	    if (m|<entry[^>]*id=\"(.*?)\"|io)
	    {
		$id = $1;
		$ERR_ID{$id} = 1;
	    }
	    printf(out_fp "%s\n", $_);
	}
    }
    close(lint_fp);
    close(out_fp);
} 


sub load_lint_line_nums
{
    my($f) = @_;
    open(lint_fp, "$f") || die "Unable to open $f"; 
    while (<lint_fp>)
    {
	chomp;
	s|^.*?:||gio;
	s|:.*||g;
	$LINT_ERR{$_} = 1;
    }
    close(lint_fp);
} 


sub get_pic_size
{
    my($f) = @_;
    my($comm);
    my($j);
    my($x, $y);    
    $comm = sprintf("identify -ping %s", $f);
    $j = `$comm`;
    #    $j =~ s|\t| |g;
    if ($j =~ m|^.*? ([0-9]+)x([0-9]+) |)
    {
	$x = $1;
	$y = $2;
    }
    #	printf("<file>%s</file><x>%s</x><y>%s</y>\n", $f, $x, $y); 
    return($y, $x); # return ht, width
}

sub lose_idmids
{
    my($e) = @_;
    $e =~ s| e:id=\"([^\"]*)\"||gi;
    $e =~ s| e:[^ >\"]*=\"[^\"]*\"||gi;
    $e =~ s| xmlns[^ >\"\=]*=\"[^\"]*\"||gi;
    #    $e =~ s| trans=| DELID=|gi;
    $e =~ s| source=| DELID=|gi;
    $e =~ s| sourceid=| DELID=|gi;
    $e =~ s| dupedid=| DELID=|gi;
    $e =~ s| dpsid=| DELID=|gi;
    $e =~ s| mediarecord=| DELID=|gi;
    $e =~ s| xmlns[^ =]*=| DELID=|gi;
    $e =~ s| original=| DELID=|gi;
    #
    $e =~ s| DELID=\"[^\"]*\"||gi;
    $e =~ s| DELID=\'[^\"]*\'||gi;    
    return $e;
}

sub reduce_idmids
{
    my($e) = @_;
    $e =~ s| e:id=| idmid=|gi;
    $e =~ s| e:targeteltid=| idmhref=|gi;
    $e =~ s| e:[^ >\"]*=\"[^\"]*\"||gi;
    $e =~ s| trans=| DELID=|gi;
    $e =~ s| source=| DELID=|gi;
    $e =~ s| sourceid=| DELID=|gi;
    $e =~ s| dupedid=| DELID=|gi;
    $e =~ s| dpsid=| DELID=|gi;
    $e =~ s| original=| DELID=|gi;
    $e =~ s| mediarecord=| DELID=|gi;
    #    $e =~ s| resource=| DELID=|gi;
    $e =~ s| dpsresource=| DELID=|gi;
    $e =~ s| xmlns[^ =]*=| DELID=|gi;
    #
    $e =~ s| DELID=\"[^\"]*\"||gi;
    $e =~ s| DELID=\'[^\"]*\'||gi;    
    $e =~ s| idmid=| e:id=|gi;
    $e =~ s| idmhref=| e:targeteltid=|gi;
    return $e;
}

sub red
{
    my($e) = @_;
    #    $e =~ s| e:id=| idmid=|gi;
    #    $e =~ s| e:targeteltid=| idmhref=|gi;
    $e =~ s| e:[^ >\"]*=\"[^\"]*\"||gi;
    $e =~ s| trans=| DELID=|gi;
    $e =~ s| source=| DELID=|gi;
    $e =~ s| sourceid=| DELID=|gi;
    $e =~ s| dupedid=| DELID=|gi;
    $e =~ s| dpsid=| DELID=|gi;
    $e =~ s| original=| DELID=|gi;
    $e =~ s| mediarecord=| DELID=|gi;
    #    $e =~ s| resource=| DELID=|gi;
    $e =~ s| dpsresource=| DELID=|gi;
    $e =~ s| xmlns[^ =]*=| DELID=|gi;
    #
    $e =~ s| DELID=\"[^\"]*\"||gi;
    $e =~ s| DELID=\'[^\"]*\'||gi;    
    $e =~ s| idmid=| e:id=|gi;
    $e =~ s| idmhref=| e:targeteltid=|gi;
    return $e;
}

sub lose_unicode_ents
{
    my($e) = @_;
    $e =~ s|&\#x2027;||gio; # &w;
    $e =~ s|&#x02C[C8];||gio;  # &[ps]
    $e =~ s|&#x2009;| |gio;
    $e =~ s|&#x2013;| |gio;
    $e =~ s|&#x2019;|\'|gio;
    $e =~ s|&\#x201C;||gio;
    $e =~ s|&\#x201D;||gio;
    $e =~ s|&\#x2022;||gio;
    $e =~ s|&\#x2026;||gio;
    $e =~ s|&\#x2027;||gio;
    $e =~ s|&\#x2122;||gio;
    $e =~ s|&\#x22C6;||gio;
    $e =~ s|&\#x25B8;||gio;
    $e =~ s|&\#x279C;||gio;
    $e =~ s|&\#x27A4;||gio;
    $e =~ s|&\#x27AA;||gio;
    $e =~ s|&\#x00B7;||gio;
    return $e;
}

sub lose_excess_h_stuff
{
    my($e) = @_;
    $e =~ s|&[swp];||gio;
    $e =~ s|&middot;||gio;
    $e =~ s|&trade;||gio;
    $e =~ s|&thinsp;||gio;
    $e =~ s|&\#x02C[C8];||gi;
    $e =~ s|&\#x2122;||gi;
    $e =~ s|&\#x2009;| |gi;
    $e =~ s|&\#x2019;|\'|gi;
    return $e;
}

sub norm_wd
{
    my($wd) = @_;
    $wd =~ s|<!--.*?-->||gio;
    $wd =~ s|&iexcl;||gio;
    $wd =~ s|\!||gio;
    $wd =~ s|&trade;||gio;
    $wd =~ s|&tm;||gio;
    $wd =~ s|&copy;||gio;
    $wd =~ s|&reg;||gio;
    $wd = &hex2letter($wd);    
    $wd =~ s|&(.)acute;|$1|gio;
    $wd =~ s|&(.)grave;|$1|gio;
    $wd =~ s|&(.)cedil;|$1|gio;
    $wd =~ s|&(.)circ;|$1|gio;
    $wd =~ s|&(.)uml;|$1|gio;
    $wd =~ s|&(.)tilde;|$1|gio;
    $wd =~ tr|A-Z|a-z|;
    $wd =~ s|[^a-z0-9]+||gio;
    return $wd;
}

sub hex2letter
{
    my($e) = @_;
    unless ($e =~ m|&\#x|)
    {
	return $e;
    }
    $e =~ s|&#x0026;|and|gi;
    $e =~ s|&\#x00C6;|AE|gio;
    $e =~ s|&\#x00C0;|A|gio;
    $e =~ s|&\#x00C1;|A|gio;
    $e =~ s|&\#x00C2;|A|gio;
    $e =~ s|&\#x00C3;|A|gio;
    $e =~ s|&\#x00C4;|A|gio;
    $e =~ s|&\#x00C5;|A|gio;
    $e =~ s|&\#x00C7;|C|gio;
    $e =~ s|&\#x00C8;|E|gio;
    $e =~ s|&\#x00C9;|E|gio;
    $e =~ s|&\#x00CA;|E|gio;
    $e =~ s|&\#x00CB;|E|gio;
    $e =~ s|&\#x00CC;|I|gio;
    $e =~ s|&\#x00CD;|I|gio;
    $e =~ s|&\#x00CE;|I|gio;
    $e =~ s|&\#x00CF;|I|gio;
    $e =~ s|&\#x00D1;|N|gio;
    $e =~ s|&\#x00D2;|O|gio;
    $e =~ s|&\#x00D3;|O|gio;
    $e =~ s|&\#x00D4;|O|gio;
    $e =~ s|&\#x00D5;|O|gio;
    $e =~ s|&\#x00D6;|O|gio;
    $e =~ s|&\#x00D8;|O|gio;
    $e =~ s|&\#x00D9;|U|gio;
    $e =~ s|&\#x00DA;|U|gio;
    $e =~ s|&\#x00DB;|U|gio;
    $e =~ s|&\#x00DC;|U|gio;
    $e =~ s|&\#x00DD;|Y|gio;
    $e =~ s|&\#x00DF;|ss|gio;
    $e =~ s|&\#x00E0;|a|gio;
    $e =~ s|&\#x00E1;|a|gio;
    $e =~ s|&\#x00E2;|a|gio;
    $e =~ s|&\#x00E3;|a|gio;
    $e =~ s|&\#x00E4;|a|gio;
    $e =~ s|&\#x00E5;|a|gio;
    $e =~ s|&\#x00E6;|ae|gio;
    $e =~ s|&\#x00E7;|c|gio;
    $e =~ s|&\#x00E8;|e|gio;
    $e =~ s|&\#x00E9;|e|gio;
    $e =~ s|&\#x00EA;|e|gio;
    $e =~ s|&\#x00EB;|e|gio;
    $e =~ s|&\#x00EC;|i|gio;
    $e =~ s|&\#x00ED;|i|gio;
    $e =~ s|&\#x00EE;|i|gio;
    $e =~ s|&\#x00EF;|i|gio;
    $e =~ s|&\#x00F1;|n|gio;
    $e =~ s|&\#x00F2;|o|gio;
    $e =~ s|&\#x00F3;|o|gio;
    $e =~ s|&\#x00F4;|o|gio;
    $e =~ s|&\#x00F5;|o|gio;
    $e =~ s|&\#x00F6;|o|gio;
    $e =~ s|&\#x00F9;|u|gio;
    $e =~ s|&\#x00FA;|u|gio;
    $e =~ s|&\#x00FB;|u|gio;
    $e =~ s|&\#x00FC;|u|gio;
    $e =~ s|&\#x00FD;|y|gio;
    $e =~ s|&\#x00FF;|y|gio;
    $e =~ s|&\#x0100;|A|gio;
    $e =~ s|&\#x0101;|a|gio;
    $e =~ s|&\#x0102;|A|gio;
    $e =~ s|&\#x0103;|a|gio;
    $e =~ s|&\#x0104;|A|gio;
    $e =~ s|&\#x0105;|a|gio;
    $e =~ s|&\#x0106;|C|gio;
    $e =~ s|&\#x0107;|c|gio;
    $e =~ s|&\#x0108;|C|gio;
    $e =~ s|&\#x0109;|c|gio;
    $e =~ s|&\#x010A;|C|gio;
    $e =~ s|&\#x010B;|c|gio;
    $e =~ s|&\#x010C;|C|gio;
    $e =~ s|&\#x010D;|c|gio;
    $e =~ s|&\#x010E;|D|gio;
    $e =~ s|&\#x010F;|d|gio;
    $e =~ s|&\#x0110;|D|gio;
    $e =~ s|&\#x0111;|d|gio;
    $e =~ s|&\#x0112;|E|gio;
    $e =~ s|&\#x0113;|e|gio;
    $e =~ s|&\#x0116;|E|gio;
    $e =~ s|&\#x0117;|e|gio;
    $e =~ s|&\#x0118;|E|gio;
    $e =~ s|&\#x0119;|e|gio;
    $e =~ s|&\#x011A;|E|gio;
    $e =~ s|&\#x011B;|e|gio;
    $e =~ s|&\#x011C;|G|gio;
    $e =~ s|&\#x011D;|g|gio;
    $e =~ s|&\#x011E;|G|gio;
    $e =~ s|&\#x011F;|g|gio;
    $e =~ s|&\#x0120;|G|gio;
    $e =~ s|&\#x0121;|g|gio;
    $e =~ s|&\#x0122;|G|gio;
    $e =~ s|&\#x0124;|H|gio;
    $e =~ s|&\#x0125;|h|gio;
    $e =~ s|&\#x0126;|H|gio;
    $e =~ s|&\#x0127;|h|gio;
    $e =~ s|&\#x0128;|I|gio;
    $e =~ s|&\#x0129;|i|gio;
    $e =~ s|&\#x012A;|I|gio;
    $e =~ s|&\#x012B;|i|gio;
    $e =~ s|&\#x012E;|I|gio;
    $e =~ s|&\#x012F;|i|gio;
    $e =~ s|&\#x0130;|I|gio;
    $e =~ s|&\#x0131;|i|gio;
    $e =~ s|&\#x0132;|IJ|gio;
    $e =~ s|&\#x0133;|ij|gio;
    $e =~ s|&\#x0134;|J|gio;
    $e =~ s|&\#x0135;|j|gio;
    $e =~ s|&\#x0136;|K|gio;
    $e =~ s|&\#x0137;|k|gio;
    $e =~ s|&\#x0139;|L|gio;
    $e =~ s|&\#x013A;|l|gio;
    $e =~ s|&\#x013B;|L|gio;
    $e =~ s|&\#x013C;|l|gio;
    $e =~ s|&\#x013D;|L|gio;
    $e =~ s|&\#x013E;|l|gio;
    $e =~ s|&\#x013F;|L|gio;
    $e =~ s|&\#x0140;|l|gio;
    $e =~ s|&\#x0141;|L|gio;
    $e =~ s|&\#x0142;|l|gio;
    $e =~ s|&\#x0143;|N|gio;
    $e =~ s|&\#x0144;|n|gio;
    $e =~ s|&\#x0145;|N|gio;
    $e =~ s|&\#x0146;|n|gio;
    $e =~ s|&\#x0147;|N|gio;
    $e =~ s|&\#x0148;|n|gio;
    $e =~ s|&\#x014C;|O|gio;
    $e =~ s|&\#x014D;|o|gio;
    $e =~ s|&\#x0152;|OE|gio;
    $e =~ s|&\#x0153;|oe|gio;
    $e =~ s|&\#x0154;|R|gio;
    $e =~ s|&\#x0155;|r|gio;
    $e =~ s|&\#x0156;|R|gio;
    $e =~ s|&\#x0157;|r|gio;
    $e =~ s|&\#x0158;|R|gio;
    $e =~ s|&\#x0159;|r|gio;
    $e =~ s|&\#x015A;|S|gio;
    $e =~ s|&\#x015B;|s|gio;
    $e =~ s|&\#x015C;|S|gio;
    $e =~ s|&\#x015D;|s|gio;
    $e =~ s|&\#x015E;|S|gio;
    $e =~ s|&\#x015F;|s|gio;
    $e =~ s|&\#x0160;|S|gio;
    $e =~ s|&\#x0161;|s|gio;
    $e =~ s|&\#x0162;|T|gio;
    $e =~ s|&\#x0163;|t|gio;
    $e =~ s|&\#x0164;|T|gio;
    $e =~ s|&\#x0165;|t|gio;
    $e =~ s|&\#x0166;|T|gio;
    $e =~ s|&\#x0167;|t|gio;
    $e =~ s|&\#x0168;|U|gio;
    $e =~ s|&\#x0169;|u|gio;
    $e =~ s|&\#x016A;|U|gio;
    $e =~ s|&\#x016B;|u|gio;
    $e =~ s|&\#x016C;|U|gio;
    $e =~ s|&\#x016D;|u|gio;
    $e =~ s|&\#x016E;|U|gio;
    $e =~ s|&\#x016F;|u|gio;
    $e =~ s|&\#x0172;|U|gio;
    $e =~ s|&\#x0173;|u|gio;
    $e =~ s|&\#x0174;|W|gio;
    $e =~ s|&\#x0175;|w|gio;
    $e =~ s|&\#x0176;|Y|gio;
    $e =~ s|&\#x0177;|y|gio;
    $e =~ s|&\#x0178;|Y|gio;
    $e =~ s|&\#x0179;|Z|gio;
    $e =~ s|&\#x017A;|z|gio;
    $e =~ s|&\#x017B;|Z|gio;
    $e =~ s|&\#x017C;|z|gio;
    $e =~ s|&\#x017D;|Z|gio;
    $e =~ s|&\#x017E;|z|gio;
    $e =~ s|&\#x01F5;|g|gio;
    $e =~ s|&\#xFB00;|ff|gio;
    $e =~ s|&\#xFB01;|fi|gio;
    $e =~ s|&\#xFB02;|fl|gio;
    $e =~ s|&\#xFB03;|ffi|gio;
    $e =~ s|&\#xFB04;|ffl|gio;
    $e =~ s|&\#x[^ ;]*;||go;
    return $e;
}

sub utf82nimpho
{
    my($e) = @_;
    $e =~ s|&#x00E6;|&amp;|gio;
    $e =~ s|&#x025C;|3|gio;
    $e =~ s|&#x0025;|6|gio;
    $e =~ s|&#x02d0;|:|gio;
    $e =~ s|&#x02D0;|;|gio;
    $e =~ s|&#x0259;|\@|gio;
    $e =~ s|&#x0251;|A|gio;
    $e =~ s|&#x00F0;|D|gio;
    $e =~ s|&#x025B;|E|gio;
    $e =~ s|&#x02D0;|F|gio;
    $e =~ s|&#x026A;|I|gio;
    $e =~ s|&#x0259;|J|gio;
    $e =~ s|&#x014B;|N|gio;
    $e =~ s|&#x0254;|O|gio;
    $e =~ s|&#x0252;|Q|gio;
    $e =~ s|&#x0283;|S|gio;
    $e =~ s|&#x0275;|T|gio;
    $e =~ s|&#x03B8;|T|gio;
    $e =~ s|&#x028A;|U|gio;
    $e =~ s|&#x028C;|V|gio;
    $e =~ s|&#x025C;|W|gio;
    $e =~ s|&#x0292;|Z|gio;
    $e =~ s|&#x0261;|g|gio;
    $e =~ s|&#x02C8;|\"|gio;
    $e =~ s|&#x02CC;|%|gio;
    $e =~ s|&#x0342;|~|gio;
    $e =~ s|&#x2011;|-|gio;
    return $e;
}

sub get_ig_phons
{
    my($e) = @_;
    $e =~ s|<(/?)ei-g([ >])|<$1i-g$2|gi;
    $e =~ s|<(/?)phon-gb([ >])|<$1i$2|gi;
    $e =~ s|<(/?)phon-us([ >])|<$1y$2|gi;
    $e =~ s|<i>|<i >|gi;
    $e =~ s|<y>|<y >|gi;
    $e =~ s|(<i [^>]*)/>|$1></i>|gi;
    $e =~ s|(<y [^>]*)/>|$1></y>|gi;
    return $e;
}

sub wdcmp
{
    my($a, $b) = @_;
    my ($bit, $wd12, $wd21);
    $a = &wdsimplify($a);
    $b = &wdsimplify($b);
    $a =~ s|z|s|g;
    $b =~ s|z|s|g;
    if ($a eq $b)
    {
	return 1;
    }
    $wd12 = sprintf("%s\t%s", $a, $b);
    $wd21 = sprintf("%s\t%s", $b, $a);
    if ($VARIANT{$wd12})
    {
	return 1;
    }
    if ($VARIANT{$wd21})
    {
	return 1;
    }

    return 0;
}

sub wdsimplify
{
    my($wd) = @_;
    $wd =~ tr|A-Z|a-z|;
    $wd =~ s|^the[ _]||;
    $wd =~ s|^xcon|con|;
    $wd =~ s|&#x00AE;||gi;
    $wd =~ s|,.*||;
    $wd =~ s|[\- _\'\.]||gi;
    return $wd;
}

sub get_xml_header
{
    my($res);
    $res = "<?xml version=\"1.0\"?><!DOCTYPE dps-data SYSTEM \"Q:\\single_dtd\\master_files\\oup_elt_dicts.dtd\"><dps-data>";
    return $res;
}

sub get_xsl_header_old
{
    my($res);
    $res = '<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:d="urn:LAENGSPA" xmlns:e="urn:IDMEE" version="1.0">&nl;<xsl:output method="xml" version="1.0" encoding="utf-8" omit-xml-declaration="yes"/>&nl;<xsl:template match="*">&nl;	<xsl:copy>&nl;		<xsl:apply-templates select="@*|node()"/>&nl;	</xsl:copy>&nl;</xsl:template>&nl;&nl;<xsl:template match="@*|text()">&nl;	<xsl:copy-of select="."/>&nl;</xsl:template>&nl;&nl;<xsl:template match="/">&nl;	<xsl:apply-templates/>&nl;</xsl:template>&nl;&nl;<xsl:template match="comment() | processing-instruction()">&nl;	<xsl:copy/>&nl;</xsl:template>&nl;&nl;';
    $res =~ s|&nl;|\n|gi;
    return $res;
}


sub get_html_header
{
    my($title, $css) = @_;
    my($res);
    if ($css =~ m|^ *$|)
    {
	$css = "single_dtd.css";
    }
    $res = sprintf("<!DOCTYPE HTML><html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><link rel=\"stylesheet\" type=\"text/css\" href=\"http://lnxcrescent/CSS/%s \"/><title>%s</title></head><body><iframe id=\"dummy\" style=\"display:none\" name=\"dummy\" height=\"1\"></iframe>", $css, $title); 
    return $res;
}

sub convert_tag_to_class
{
    my($e, $tagname) = @_;
    my($res, $eid);
    my($bit);
    my(@BITS);
    $e = restructure::delabel($e);
    $e =~ s|(<$tagname)>|$1 >|gi;
    $e =~ s|<$tagname .*?>|<span class=\"$tagname\">|gi;
    $e =~ s|</$tagname>|</span>|gi;
    return $e;
}

sub convert_tag_to_div
{
    my($e, $tagname) = @_;
    my($res, $eid);
    my($bit);
    my(@BITS);
    $e = restructure::delabel($e);
    $e =~ s|(<$tagname)>|$1 >|gi;
    $e =~ s|<$tagname .*?>|<div class=\"$tagname\">|gi;
    $e =~ s|</$tagname>|</div>|gi;
    return $e;
}

sub load_posdefs
{
    $POS{"adjective"} = "adj";
    $POS{"adverb"} = "adv";
    $POS{"day"} = "n";
    $POS{"indefa"} = "det";
    $POS{"indef art"} = "det";
    $POS{"day"} = "n";
    $POS{"defa"} = "det";
    $POS{"determiner"} = "det";
    $POS{"exclamation"} = "exclam";
    $POS{"idiom"} = "idm";
    $POS{"modal verb"} = "v";
    $POS{"month"} = "n";
    $POS{"noun"} = "n";
    $POS{"noun / verb"} = "n";
    $POS{"num"} = "adj";
    $POS{"number"} = "adj";
    $POS{"ordinal"} = "adj";
    $POS{"preposition"} = "prep";
    $POS{"pronoun"} = "pron";
    $POS{"verb"} = "v";
    $POS{"pv"} = "v";
    $POS{"vt"} = "v";
    $POS{"vi"} = "v";
    $POS{"vmodal"} = "v";
}

sub trans_pos
{
    my($e) = @_;
    my($res, $eid);	
    $e =~ s| *$||;
    $res = $POS{$e};
    if ($res =~ m|[a-z]|)
    {
	return $res;
    }
    return $e;
}


sub get_idiom_sortkey
{
    my($e) = @_;
    my($res, $eid);	
    $key = $e;
    $key =~ s|<[^>]*>||go;
    $key =~ s|&#x2019;|\'|gi;
    $key =~ s|&ob;|/|go;
    $key =~ s|&[^;]*;||go;
    $key =~ s|,? etc\.?||go;
    #
    # Sorting rules from ALD5
    $key =~ s|\(.*?\)||go;
    $key =~ s|/[^ ]*||go;
    $key = sprintf(" %s ", $key);
    $key =~ tr|A-Z|a-z|;
    $key =~ s|([^a-z])sb\'s([^a-z])|$1 $2|;
    $key =~ s|([^a-z])sth\'s([^a-z])|$1 $2|;
    $key =~ s|([^a-z])sb([^a-z])|$1 $2|;
    $key =~ s|([^a-z])sth([^a-z])|$1 $2|;
    $key =~ s/\b(?:one's|my|your|his|her|its|their)\b//go;
    $key =~ s/\b(?:the|an?)\b//go;
    $key =~ s|\W+||go;
    return($key);
}

sub numerical_hash_sort
{
    my($HASH, $ARR) = @_;
    my($res, $eid, $key, $ct);	    
    foreach $key (sort { $HASH->{$a} <=> $HASH->{$b} } keys %$HASH) {
	$ARR->[$ct++] = $key;
    }
}

sub print_dtd_hdr
{
    printf("<?xml version=\"1.0\"?>\n<!DOCTYPE dps-data SYSTEM \"C:\\Program Files (x86)\\Blast Radius\\XMetaL 4.6\\Author\\Rules\\oup_elt_dicts.dtd\">\n<dps-data>\n");
}

sub hex_non_ascii
{
    my($e) = @_;
    my($res, $eid);	
    $e =~ s|([\!-\} ]+)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gi){
	    $bit =~ s/(.)/sprintf("&\#x%x;",ord($1))/eg;
	}
	$res .= $bit;
    }
    return $res;
}

&def_prevletter;
&def_nextletter;
&define_wdlist_links;
1;
