#!/usr/bin/perl
use Getopt::Std;
use autodie qw(:all);
use utf8;
use strict;
our ($LOG, $LOAD, $opt_f, $opt_u, $opt_D, $opt_I, $opt_O, $UTF8, %USED, @PTAGS, %LBLS);
require "./utils.pl";
require "./restructure.pl";
#require "/NEWdata/dicts/generic/progs/xsl_lib_fk.pl";
$LOG = 1;
$LOAD = 0;
$UTF8 = 1;
$, = ' ';               # set output field separator
$\ = "\n";              # set output record separator
#undef $/; # read in the whole file at once
&main;

sub main
{
    getopts('uf:L:IOD');
    &usage if ($opt_u);
    my($e, $res, $bit);
    my(@BITS);
    #   $opt_L = ""; # name of file for the log_fp output to go to
    use open qw(:std :utf8);
    &open_debug_files;
    if ($opt_D)
    {
	binmode DB::OUT,":utf8";
    }
    @PTAGS = ("se1", "subEntry");
    if ($LOAD){&load_file($opt_f);}
  line:    
    while (<>){
	chomp;       # strip record separator
	s|||g;
	if ($opt_I){printf(bugin_fp "%s\n", $_);}
	#	unless (m|<entry|){print $_; next line;}
	# s|<!--.*?-->||gio;
	#	next line if (m|<entry[^>]*del=\"y|io);
	#	next line if (m|<entry[^>]*sup=\"y|io);
	# $h = &get_hex_h($_, "hex", 1); # the 1 says to remove stress etc
	# $eid = &get_tag_attval($_, "entry", "eid");
	# $e_eid = &get_dps_entry_id($_);
	# $_ = &reduce_idmids($_);
	# s|£|&\#x00A3;|g;
	$_ = restructure::delabel($_);	
	# $tagname = restructure::get_tagname($bit);    
	$_ = &inherit_labels($_);
	$_ = &inherit_pos($_);
	$_ = &add_dr_forms($_);
	$_ = &add_hws($_);
	$_ = &combine_labels($_);
	$_ = &add_sense_numbers($_);
	print $_;
	if ($opt_O){printf(bugout_fp "%s\n", $_);}
    }
    &close_debug_files;
}

sub add_sense_numbers
{
    my($e) = @_;
    my($res, $eid);	
    my($bit, $res);
    my(@BITS);
    my $defct = ($e =~ s|(</df>)|\1|gi);
    $e = restructure::set_tag_attval($e, "msDict", "defct", $defct);
    $e =~ s|(<se1[ >].*?</se1>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $se2ct = ($bit =~ s|(</se2>)|\1|gi);
	    $bit = restructure::set_tag_attval($bit, "msDict", "se2ct", $se2ct); 
	    $bit = &add_sensenums($bit);
	}
	$res .= $bit;
    }    
    return $res;
}

sub add_sensenums
{
    my($e) = @_;
    my($res, $eid);	
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<se2[ >].*?</se2>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $sensenum = restructure::get_tag_attval($bit, "se2", "num"); 
	    $bit = restructure::set_tag_attval($bit, "msDict", "sensenum", $sensenum);
	}
	$res .= $bit;
    }    
    return $res;
}

sub combine_labels
{
    my($e) = @_;
    my($res, $eid);	
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<msDict[ >].*?</msDict>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $labels = &get_labels($bit);
	    $bit =~ s|(</msDict>)|$labels\1|;
	}
	$res .= $bit;
    }    
    return $res;
}

sub get_labels
{
    my($e) = @_;
    my($res, $eid);	
    my($bit, $res);
    my(@BITS);
    undef %LBLS;
    $e =~ s|(<lg[ >].*?</lg>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $label = restructure::get_tag_contents($bit, "lg");
	    $label =~ s|<.*?>||gi;
	    $LBLS{$label} = 1;
	}
    }    
    foreach my $label (keys %LBLS)
    {
	$res = sprintf("%s%s, ", $res, $label); 
    }
    $res =~ s|, *$||;
    unless ($res =~ m|^ *$|)
    {
	$res = sprintf("<labels>%s</labels>", $res);
	printf(log_fp "%s\n", $res); 
    }
    return $res;
}

sub add_hws
{
    my($e) = @_;
    my($res, $eid, $hw);	
    my(@BITS);
    if ($e =~ m|<hw|)
    {
	$hw = restructure::get_tag_attval($e, "hw", "normalized"); 
	if ($hw =~ m|^ *$|)
	{
	    $hw = restructure::get_tag_contents($e, "hw"); 
	}
    }
    $e =~ s|(<msDict[ >].*?>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $form = restructure::get_tag_attval($bit, "msDict", "form");
	    if ($form =~ m|^ *$|)
	    {
		$bit = restructure::set_tag_attval($bit, "msDict", "form", $hw); 
	    }
	}
	$res .= $bit;
    }    
    return $res;
}



sub add_dr_forms
{
    my($e) = @_;
    my($res, $eid);	
    $e =~ s|(<subEntryBlock[ >].*?</subEntryBlock>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $e);
    $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ m|</l>|)
	    {
		$bit = &add_drs_to_defs($bit);
	    }
	}
	$res .= $bit;
    }	
    return $res;
}

sub add_drs_to_defs
{
    my($e) = @_;
    my($res, $eid);	
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<subEntry[ >].*?</subEntry>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $form = &get_ls($bit);
	    $bit =~ s|<msDict |<msDict form=\"$form\" |g;	    
	}
	$res .= $bit;
    }    
    return $res;
}

sub get_ls
{
    my($e) = @_;
    my($res, $eid);	
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<l[ >].*?</l>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $form = restructure::get_tag_contents($bit, "l"); 
	    $res = sprintf("%s%s&sep; ", $res, $form); 
	}
    }
    $res =~ s|&sep; *$||;
    $res =~ s/&sep;/\|/g;
    return $res;
}

sub inherit_pos
{
    my($e) = @_;
    my($res, $eid);	
    $_ =~ s|(<se1[ >].*?</se1>)|&split;&fk;$1&split;|gi;
    my @BITS = split(/&split;/, $_);
    $res = "";
    foreach my $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $pos = restructure::get_tag_contents($bit, "pos");
	    $bit =~ s|<msDict |<msDict pos=\"$pos\" |gi;
	}
	$res .= $bit;
    }    
    return $res;
}

sub inherit_labels
{
    my($e) = @_;
    my($res, $eid);	
    $e = restructure::add_levels_info($e, "e");
    my $ptag;
    foreach $ptag (@PTAGS)
    {
	$e = &inherit_tags($e, $ptag, "lg");
    }
    $e = &inherit_hw_tags($e);
    $e =~ s| level=\".*?\"||gi;
    return $e;
}




sub inherit_hw_tags
{
    my($e, $ptag, $tag) = @_;
    my($res, $eid);	
    my($bit, $res);
    my(@BITS);
    if ($e =~ m|(<hg[ >].*?</hg>)|)
    {
	my $hg = $1;
	my $labels = &get_tag_at_level($hg, "lg", 2);
	unless ($labels =~ m|^ *$|)
	{
	    $e = &inherit_to_msdict($e, $labels);
	}		
    }
    return $e;
}


sub inherit_tags
{
    my($e, $ptag, $tag) = @_;
    my($res, $eid);	
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<$ptag[ >].*?</$ptag>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ m|<$tag |)
	    {
		my $level = restructure::get_tag_attval($bit, $ptag, "level");
		my $tlevel = $level + 1;
		my $labels = &get_tag_at_level($bit, $tag, $tlevel);		
		unless ($labels =~ m|^ *$|)
		{
		    $bit = &inherit_to_msdict($bit, $labels);
		}
		#		print $bit;
		#		print $labels;
	    }
	}
	$res .= $bit;
    }    
    return $res;
}

sub inherit_to_msdict
{
    my($e, $labels) = @_;
    my($res, $eid);
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<msDict[ >].*?</msDict>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    unless ($bit =~ m|<plabels|)
	    {
		$bit =~ s|</msDict>|<plabels></plabels></msDict>|;
	    }
	    $bit =~ s|(</plabels>)|$labels\1|;
	}
	$res .= $bit;
    }    
    return $res;
}

sub get_tag_at_level
{
    my($e, $tag, $tlevel) = @_;
    my($res, $eid);	
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<$tag[ >].*?</$tag>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    my $level = restructure::get_tag_attval($bit, $tag, "level");
	    if ($level == $tlevel)
	    {
		$bit =~ s| level=\".*?\"||g;
		$bit =~ s| e:[^= ]*=\".*?\"||g;
		$res .= $bit;
	    }
	}
    }    
    return $res;
}


sub usage
{
    printf(STDERR "USAGE: $0 -u \n"); 
    printf(STDERR "\t-u:\tDisplay usage\n"); 
    #    printf(STDERR "\t-x:\t\n"); 
    exit;
}


sub load_file
{
    my($f) = @_;
    my ($res, $bit, $info);
    my @BITS;
    open(in_fp, "$f") || die "Unable to open $f"; 
    if ($UTF8){
	binmode(in_fp, ":utf8");
    }
    while (<in_fp>){
	chomp;
	s|||g;
	# my ($eid, $info) = split(/\t/);
	# $W{$_} = 1;
    }
    close(in_fp);
} 
