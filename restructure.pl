#!/usr/local/bin/perl
#
# Input = 
# Result = 
# $Id: restructure.pl,v 1.6 2019/09/24 14:41:47 keenanf Exp keenanf $
# $Log: restructure.pl,v $
# Revision 1.6  2019/09/24 14:41:47  keenanf
# *** empty log message ***
#
# Revision 1.5  2019/02/28 08:50:04  keenanf
# *** empty log message ***
#
# Revision 1.4  2017/10/16 13:32:57  keenanf
# *** empty log message ***
#
# Revision 1.3  2017/01/27 11:53:02  keenanf
# *** empty log message ***
#
# Revision 1.2  2014/12/05 14:24:22  keenanf
# *** empty log message ***
#
# Revision 1.1  2014/11/30 10:05:52  keenanf
# Initial revision
#
#
package restructure;
#require "/disk1/home/keenanf/perl/utils.pl";
#
#
# Load this with
# require "/NEWdata/dicts/generic/progs/restructure.pl";
#
#Calls
# $e = restructure::delabel($e);
# $e = restructure::strip_dps($e);
# $tagname = restructure::get_tagname($bit);    
# $e = restructure::add_group_tag_outsidegroup($e, "TAG", "GROUP_TAG", "EXCLUDE_GROUP");
# $e = restructure::add_group_tag($e, "TAG", "GROUP_TAG", "NO_COMBINE"); # if NO_COMBINE == 1 won't group them together
# $e = restructure::add_xgs($e, $dict);
# $e = restructure::add_xgss($e);
# $e = restructure::add_bilingual_xgs($e);
# $e = restructure::add_transgss($e);
# $e = restructure::add_bfsgs($e);
# $e = restructure::add_cpdgs($e);
# $e = restructure::add_idsgs($e);
# $e = restructure::add_idmgs($e);
# $e = restructure::add_ifgs($e);
# $e = restructure::add_pvsgs($e);
# $e = restructure::add_disg($e);
# $e = restructure::add_vsgs($e);
# $e = restructure::add_posg($e, $dict);
# $e = restructure::add_posgs($e);
# $e = restructure::add_gramgs($e);
# $e = restructure::add_sngs($e);
# $e = restructure::add_bfgs($e);
# $e = restructure::do_ndvs($e);
# $e = restructure::expand_tildes($e);
# $e = restructure::mark_empty_phon_groups($e);
# $e = restructure::add_xrgs($e);
# $e = restructure::eb_in_if($e);
# $e = restructure::eb_in_d($e);
# $e = restructure::eb_in_x($e);
# $e = restructure::set_psg_attval($e, "TAGNAME", "ATTVAL");
# $e = restructure::do_prac_pron($e);
# $e = restructure::put_audio_on_ig_tags($e);
# $e = restructure::get_spoken_examples($e);
# $e = restructure::add_levels_info($e, "entry");
# $e = restructure::add_zp_to_tag_following_unbox($e);
# $e = restructure::tag_rename($e, "OLDTAGNAME", "NEWTAGNAME", "TOFIX", "PSGCOMMENT");
# $e = restructure::rename_tag_in_tag($e, $container, $oldtag, $newtag);
# $e = restructure::rename_tag_with_attrib($e, "OLDTAGNAME", "ATTNAME", "ATTVAL", "NEWTAGNAME");
# $e = restructure::del_tag_with_attrib($e, "TAGNAME", "ATTNAME", "ATTVAL");
# $e = restructure::renumber_ngs($e);
# $e = restructure::attrib_rename($e, "TAGNAME", "OLDATTNAME", "NEWATTNAME");
# $e = restructure::tag_delete($e, "TAGNAME");
# $e = restructure::tag_delete_empty($e, "TAGNAME");
# $e = restructure::lose_tag($e, "TAGNAME"); # lose the tags but not the contents
# $e = restructure::catname($e, "TAGNAME", "OLDATTNAME", "NEWATTNAME");
# $e = restructure::catval($e, "TAGNAME", "ATTNAME", "OLDATTVAL", "NEWATTVAL", "TOFIX", "PSGCOMMENT");
# $e = restructure::del_attrib($e, "TAGNAME", "ATTNAME", "TOFIX", "PSGCOMMENT");
# $e = restructure::setqattrib($e, "TAGNAME", "ATTNAME", "OLDATTVAL", "QVAL", "NEWATTVAL", "TOFIX", "PSGCOMMENT");
# $e = restructure::get_attval($bit, "ATTNAME");    
# $e = restructure::get_tag_attval($bit, "TAGNAME", "ATTNAME");    
# $e = restructure::set_tag_attval($bit, "TAGNAME", "ATTNAME", "ATTVAL");    
# $e = restructure::set_first_tag_attval($bit, "TAGNAME", "ATTNAME", "ATTVAL");    
# $e = restructure::hide_comments($e);
# $e = restructure::neutralise_comments($e);
# $e = restructure::restore_comments($e);
# $e = restructure::group_pgs($e);
# $e = restructure::get_tag_contents($bit, "TAGNAME");    
# $e = restructure::get_tag($bit, "TAGNAME");    
# $e = restructure::reduce_prongs($e);
# $e = restructure::add_missing_attvals($e, $tagname, $attname, $genericval);
# $e = restructure::clean_commments($e);
# $e = restructure::move_attrib($e, "GROUPTAG", "FROMTAG", "TOTAG", "ATTNAME");
#
# MOVER FUNCTIONS - always give tag names in the order they start
#    $e = restructure::move_back_into($e, "PRETAG", "POSTTAG");
#    $e = restructure::move_back_out_of($e, "PRETAG", "POSTTAG");
#    $e = restructure::move_forward_into($e, "PRETAG", "POSTTAG");
#    $e = restructure::move_forward_out_of($e, "PRETAG", "POSTTAG");
#    ($e, $extract) = restructure::extract_tag($e, $tagname);
#    $e = restructure::expand_sb_sth($e);
##

sub delabel
{
    my($e, $tag) = @_;
    my($bit, $res);
    my(@BITS);
    unless ($e =~ m|/>|)
    {
	return $e;
    }
    unless ($tag =~ m|^ *$|)
    {
	$e =~ s|(<$tag [^>]*)/>|\1></$tag>|gi;
	return $e;
    }
    $e =~ s|(<[a-z][^>]*/>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s| */>| />|;
	    $bit =~ s|<(([^ ]*) .*)/>|<$1></$2>|;
	}
	$res .= $bit;
    }
    
    return $res;
}


sub strip_dps
{
    my($e) = @_;
    $e =~ s| e:target| £:target|gi;
    $e =~ s| e:[^ =]*=\".*?\"| |gi;
    $e =~ s| xmlns[^ =]*=\".*?\"| |gi;
    $e =~ s| £:target| e:target|gi;
    $e =~ s| +| |g;
    return $e;
}


sub add_group_tag_outsidegroup
{
    my($e) = @_;
    my($e, $tag, $grouptag, $not_in_group) = @_;
    if ($not_in_group =~ m|^ *$|)
    {
	$not_in_group = $grouptag;
    }
    my($res, $eid);	
    unless ($e =~ m|<$tag[> ]|i)
    {
	return($e);	
    }
    if ($e =~ m|<$tag [^>]*/>|)
    {
	$e = restructure::delabel($e);
    }
    my(@BITS);
    $e =~ s|(<$not_in_group[ >].*?</$not_in_group>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gi){
	    $bit = &add_group_tag($bit, $tag, $grouptag);
	}
	$res .= $bit;
    }    
    return $res;
}


sub add_group_tag
{
    my($e, $tag, $grouptag, $no_combine) = @_;
    my($bit, $res, $eid);	
    unless ($e =~ m|<$tag[> ]|i)
    {
	return($e);	
    }
    my(@BITS);
    $e =~ s|(<$grouptag[ >].*?</$grouptag>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gi){
	    $bit =~ s|(<$tag[ >].*?</$tag>)|<FKGRP >$1</FKGRP>|gi;
	}
	$res .= $bit;
    }    
    unless ($no_combine)
    {
	$res =~ s|</FKGRP> *<FKGRP >||gi;
    }
    $res =~ s|<FKGRP |<$grouptag |gi;
    $res =~ s|</FKGRP>|</$grouptag>|gi;
    return $res;
}


sub add_xrgs_safe
{
    my($e) = @_;
    my($bit, $res, $xt, $xt1, $xt2);
    my(@BITS);
    $e =~ s|(<xr[ >].*?</xr>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS)
    {
	if ($bit =~ s|&fk;||gio)
	{
	    $xt = "";
	    if ($bit =~ m| xt=\"(.*?)\"|i) # 
	    {
		$xt = $1;
	    }		# 
	    else {
		$xt = "see";
	    }
	    $bit = sprintf("<xr-g xt=\"%s\">%s</xr-g xt=\"%s\">", $xt, $bit, $xt);
	}
	$res .= $bit;
    }
    $e = $res;
    $e =~ s|(</xr-g[^>]*><xr-g[^>]*>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS)
    {
	if ($bit =~ s|&fk;||gio)
	{
	    if ($bit =~ m|xt=\"(.*?)\".*xt=\"(.*?)\"|i)
	    {
		$xt1 = $1;
		$xt2 = $2;
		if ($xt1 eq $xt2)
		{
		    $bit = "";
		}
	    }
	}
	$res .= $bit;
    }
    $res =~ s|</xr-g [^>]*>|</xr-g>|gio;
    return $res;
}


sub add_xrgs
{
    my($e) = @_;
    my($bit, $xt, $xt1, $xt2, $res);
    my(@BITS);
    # avoid adding xr-gs if already there
    $e =~ s|(<xr-g[ >].*?</xr-g>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	}
	else
	{
	    if ($bit =~ m|<xr|)
	    {
		$bit = &add_xrgs_to_tags($bit);
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_xrgs_to_tags
{
    my($e) = @_;
    my($bit, $xt, $xt1, $xt2, $res);
    my(@BITS);
    $e =~ s|(<xr[ >].*?</xr>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS)
    {
	if ($bit =~ s|&fk;||gio)
	{
	    $xt = "";
	    if ($bit =~ m| xt=\"(.*?)\"|i) # 
	    {
		$xt = $1;
	    }
	    else {
		$xt = "see";
	    }
	    $bit = sprintf("<xr-g xt=\"%s\">%s</xr-g xt=\"%s\">", $xt, $bit, $xt);
	}
	$res .= $bit;
    }
    $e = $res;
    $e =~ s|(</xr-g[^>]*><xr-g[^>]*>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS)
    {
	if ($bit =~ s|&fk;||gio)
	{
	    if ($bit =~ m|xt=\"(.*?)\".*xt=\"(.*?)\"|i)
	    {
		$xt1 = $1;
		$xt2 = $2;
		if ($xt1 eq $xt2)
		{
		    $bit = "";
		}
	    }
	}
	$res .= $bit;
    }
    $res =~ s|</xr-g [^>]*>|</xr-g>|gio;
    return $res;
}

sub add_bfsgs
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <idm-gs tags then they are not repeated
    my(@BITS);
    $e =~ s|(<bf-gs[ >].*?</bf-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<bf-g[ >]|)
	    {
		$bit =~ s|(<bf-g[ >].*?</bf-g>)|<bf-gs>$1</bf-gs>|gio;
		$bit =~ s|</bf-gs><bf-gs[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_idmgs
{
    my($e) = @_;
    my($res, $eid);	
    $e =~ s|(<idm-gs[ >].*?</idm-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if (($bit =~ m|<idm-g[ >]|) || ($bit =~ m|<idmxr-gs[ >]|))
	    {
		$bit =~ s|(<id-g[ >].*?</id-g>)|<idm-gs>$1</idm-gs>|gio;
		$bit =~ s|(<idm-g[ >].*?</idm-g>)|<idm-gs>$1</idm-gs>|gio;
		$bit =~ s|(<idmxr-gs[ >].*?</idmxr-gs>)|<idm-gs>$1</idm-gs>|gio;
		$bit =~ s|</idm-gs><idm-gs[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_idsgs
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <idm-gs tags then they are not repeated
    my(@BITS);
    $e =~ s|(<idm-gs[ >].*?</idm-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if (($bit =~ m|<idm?-g[ >]|) || ($bit =~ m|<idmxr-gs[ >]|))
	    {
		$bit =~ s|(<id-g[ >].*?</id-g>)|<idm-gs>$1</idm-gs>|gio;
		$bit =~ s|(<idm-g[ >].*?</idm-g>)|<idm-gs>$1</idm-gs>|gio;
		$bit =~ s|(<idmxr-gs[ >].*?</idmxr-gs>)|<idm-gs>$1</idm-gs>|gio;
		$bit =~ s|</idm-gs><idm-gs[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_ifgs
{
    my($e) = @_;
    my($bit, $res);
    # have ifs-g but nothing to separate the if's if multiples
    my(@BITS);
    $e =~ s|(<il [^>]*)> *</il>|$1/>|gi;
    $e =~ s|(<ifs-g[ >].*?</ifs-g>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    unless ($bit =~ m|<if-g|i)
	    {
		$bit =~ s|(<if[ >].*?</if>)|<if-g >$1</if-g>|gi;
		while ($bit =~ s|(</if-g>) *(<e?i-g[ >].*?</e?i-g>)|$2$1|gi)
		{}
		while ($bit =~ s|(<il [^>]*/>) *(<if-g[^>]*>)|$2$1|gi)
		{}
	    }
	}
	$res .= $bit;
    }
    if ($res =~ m|<il |i)
    {
	$res = &do_ifg_ils($res);
    }
    return $res;
}

sub do_ifg_ils
{
    my($e) = @_;
    my($res, $eid);	
    my($bit, $il);
    my(@BITS);
    $e =~ s|(<if-g[ >].*?</if-g>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ m|<il |i)
	    {
		$il = restructure::get_tag_attval($bit, "il", "il"); 
		$bit = restructure::set_tag_attval($bit, "if-g", "form", $il); 
		$bit =~ s|<il [^>]*>||g;
		$bit =~ s|</il *>||g;
	    }
	}
	$res .= $bit;
    }    
    return $res;
}


sub del_tag_with_attrib
{
    my($e, $tagname, $attname, $attval) = @_;
    my($res, $eid);	
    my($bit, $val);
    my(@BITS);
    $e =~ s|(<$tagname[ >].*?</$tagname>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $val = restructure::get_tag_attval($bit, $tagname, $attname); 
	    if ($val eq $attval)
	    {
		$bit = "";
	    }
	}
	$res .= $bit;
    }    
    return $res;
}

sub rename_tag_with_attrib
{
    my($e, $tagname, $attname, $attval, $newtagname) = @_;
    my($res, $eid);	
    my($bit, $val);
    my(@BITS);
    $e =~ s|(<$tagname[ >].*?</$tagname>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $val = restructure::get_tag_attval($bit, $tagname, $attname); 
	    if ($val eq $attval)
	    {
		$bit = restructure::tag_rename($bit, $tagname, $newtagname);
	    }
	}
	$res .= $bit;
    }    
    return $res;
}

sub add_sngs
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <sn-g tags then they are not repeated
    my(@BITS);
    $e =~ s|(<sn-gs[ >].*?</sn-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<sn-g[ >]|)
	    {
		$bit =~ s|(<sn-g[ >].*?</sn-g>)|<sn-gs>$1</sn-gs>|gio;
		$bit =~ s|</sn-gs><sn-gs[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_bfgs
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <bf-g tags then they are not repeated
    my(@BITS);
    $e =~ s|(<bf-gs[ >].*?</bf-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<bf-g[ >]|)
	    {
		$bit =~ s|(<bf-g[ >].*?</bf-g>)|<bf-gs>$1</bf-gs>|gio;
		$bit =~ s|</bf-gs><bf-gs[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}


sub add_xgss
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <x-g tags then they are not repeated
    my(@BITS);
    $e =~ s|(<x-gs[ >].*?</x-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<x-g[ >]|)
	    {
		$bit =~ s|(<x-g[ >].*?</x-g>)|<x-gs>$1</x-gs>|gio;
		$bit =~ s|</x-gs><x-gs[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_disg
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <dis-g tags then they are not repeated
    my(@BITS);
    $e =~ s|(<dis-g[ >].*?</dis-g>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    $bit =~ s|(<ds [^>]*> *</ds>)|<dis-g>$1</dis-g>|gio;
	    $bit =~ s|(<dlf [^>]*> *</dlf>)|<dis-g>$1</dis-g>|gio;
	    $bit =~ s|(<dtxt[ >].*?</dtxt>)|<dis-g>$1</dis-g>|gio;
	    $bit =~ s|(<dacadv[ >].*?</dacadv>)|<dis-g>$1</dis-g>|gio;
	    $bit =~ s|(<dnov[ >].*?</dnov>)|<dis-g>$1</dis-g>|gio;
	    $bit =~ s|(<dnca[ >].*?</dnca>)|<dis-g>$1</dis-g>|gio;
	    $bit =~ s|(<dncn[ >].*?</dncn>)|<dis-g>$1</dis-g>|gio;
	    $bit =~ s|(<dnsv[ >].*?</dnsv>)|<dis-g>$1</dis-g>|gio;
	    $bit =~ s|(<dsyn[ >].*?</dsyn>)|<dis-g>$1</dis-g>|gio;
	    $bit =~ s|(<dvcadv[ >].*?</dvcadv>)|<dis-g>$1</dis-g>|gio;
	    $bit =~ s|</dis-g><dis-g[^>]*>||gio;
	}
	$res .= $bit;
    }
    return $res;
}


sub add_transgss
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <trans-g tags then they are not repeated
    my(@BITS);
    $e =~ s|(<trans-gs[ >].*?</trans-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<trans-g[ >]|)
	    {
		$bit =~ s|(<trans-g[ >].*?</trans-g>)|<trans-gs>$1</trans-gs>|gio;
		$bit =~ s|</trans-gs><trans-gs[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_cpdgs
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <cpd-g tags then they are not repeated
    my(@BITS);
    $e =~ s|(<cpd-gs[ >].*?</cpd-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<cpd-g[ >]|)
	    {
		$bit =~ s|(<cpd-g[ >].*?</cpd-g>)|<cpd-gs>$1</cpd-gs>|gio;
		$bit =~ s|</cpd-gs><cpd-gs[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}


sub add_posgs
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <pos-g tags then they are not repeated
    my(@BITS);
    $e =~ s|(<pos-g[ >].*?</pos-g>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<pos[ >]|)
	    {
		$bit =~ s|(<pos[ >].*?</pos>)|<pos-g>$1</pos-g>|gio;
		$bit =~ s|</pos-g><pos-g[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}



sub add_gramgs
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <gram-g tags then they are not repeated
    my(@BITS);
    $e =~ s|(<gram-g[ >].*?</gram-g>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<gram[ >]|)
	    {
		$bit =~ s|(<gram[ >].*?</gram>)|<gram-g>$1</gram-g>|gio;
		$bit =~ s|</gram-g><gram-g[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}


sub add_drgs
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <idm-gs tags then they are not repeated
    my(@BITS);
    $e =~ s|(<dr-gs[ >].*?</dr-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<dr-g[ >]|)
	    {
		$bit =~ s|(<dr-g[ >].*?</dr-g>)|<dr-gs>$1</dr-gs>|gio;
		$bit =~ s|</dr-gs><dr-gs[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}


sub add_vsgs
{
    my($e) = @_;
    my($bit, $res, $type);
    # change the logic so that if there are already <v-gs tags then they are not repeated
    my(@BITS);
    $e =~ s|(<v-gs[ >].*?</v-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<v-g[ >]|)
	    {
		$bit =~ s|(<v-g[ >].*?</v-g>)|<v-gs>$1</v-gs>|gio;
		$bit =~ s|</v-gs><v-gs[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    $e = $res;
    $e =~ s|(<v-gs[ >].*?</v-gs>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ m| type=\"(.*?)\"|i)
	    {
		$type = $1;
		$bit = restructure::set_tag_attval($bit, "v-gs", "type", $type); 
	    }
	}
	$res .= $bit;
    }    
    return $res;
}


sub add_pvsgs
{
    my($e) = @_;
    my($bit, $res);
    # change the logic so that if there are already <idm-gs tags then they are not repeated
    my(@BITS);
    $e = restructure::tag_rename($e, "pvs-g", "pv-gs");
    $e = restructure::add_group_tag($e, "pvp-g", "pv-gs");
    $e =~ s|</pv-gs><pv-gs [^>]*>||gi;
    $e =~ s|(<pv-gs[ >].*?</pv-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if (($bit =~ m|<pv-g[ >]|) || ($bit =~ m|<pvxr-gs[ >]|))
	    {
		$bit =~ s|(<pv-g[ >].*?</pv-g>)|<pv-gs>$1</pv-gs>|gio;
		$bit =~ s|(<pvxr-gs[ >].*?</pvxr-gs>)|<pv-gs>$1</pv-gs>|gio;
		$bit =~ s|</pv-gs><pv-gs[^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub mark_empty_phon_groups
{
    my($e) = @_;
    my($bit, $res, $cp);
    my(@BITS);
    $e =~ s|(<i-g[ >].*?</i-g>)|&split;&fk;$1&split;|gio;
    $e =~ s|(<ei-g[ >].*?</ei-g>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    $cp = $bit;
	    $cp =~ s| sup=\"y| display=\"n|gio;
	    $cp =~ s|<(/?)phon-gb|<$1phon|gio;
	    $cp =~ s|<(/?)phon-us|<$1phon|gio;
	    $cp =~ s|<(/?)i([ >])|<$1phon$2|gio;
	    $cp =~ s|<phon[^>]* display=\"n\".*?</phon>||gio;
	    $cp =~ s|<phon[^>]*> *</phon>||gio;
	    unless ($cp =~ m|<phon|i)
	    {
		unless ($bit =~ m|<e?i-g[^>]*slashes=\"|i)
		{
		    $bit =~ s|(<e?i-g[^>]*)>|$1 slashes=\"n\">|i;
		}
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_xgs
{
    my($e, $dict) = @_;
    my($bit, $res);
    # change the logic so that if there are already <x-g tags then they are not repeated
    my(@BITS);
    $e =~ s|(<wx-g[ >].*?</wx-g *>)|&split;&fk;$1&split;|gio;
    $e =~ s|(<x-g[ >].*?</x-g *>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<x[ >]|)
	    {
		$bit = &add_xg_to_dat($bit, $dict);
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_bilingual_xgs
{
    my($e) = @_;
    my($res, $eid);	
    my($bit);
    my(@BITS);
    $e =~ s|<x>|<x >|gi;
    $e =~ s|(<x-g[ >].*?</x-g>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gi){
	    # x-g will be pairs of <x> and <tx>
	    $bit =~ s|</tx>|</tx>£|gi;
	    $bit =~ s|(<x[ >])|£$1|gi;
	    $bit =~ s|(<x [^£]*</x> *<tx[^£]*</tx>)|<x-g >$1</x-g>|gi;
	    $bit =~ s|£||g;
	}
	$res .= $bit;
    }    
    return $res;
}

sub add_xg_to_dat
{
    my($e, $dict) = @_;
    $e =~ s|£|&\#x00A3;|gio;
    $e =~ s|(<x[ >].*?</x>)|<x-g-tmp>$1</x-g-tmp>|gio;
    $e =~ s|(<tx[ >].*?</tx>)|<x-g-tmp>$1</x-g-tmp>|gio;
    $e =~ s|(<[grsg] [^>]*/>)|<label-g>$1</label-g>|gio;
    if ($dict =~ m|oad|i)
    {
	$e =~ s|(<u[ >])|£$1|gi;
	$e =~ s|(<u[ >][^£]*</u>) *(<x [^>]*>)|<label-g>$1</label-g>$2|gi;
	$e =~ s|£||g;
	$e =~ s|(<pt [^>]*/>)|<label-g>$1</label-g>|gio;
	$e =~ s|(<gr [^>]*/>)|<label-g>$1</label-g>|gio;
    }
    if ($dict =~ m|awp2e|i)
    {
	$e =~ s|<(/?)(pt[ >])|<$1£$2|gi;
	$e =~ s|<(/?)([ru][ >])|<$1£$2|gi;
	$e =~ s|<(/?)(gr[ >])|<$1£$2|gi;
	$e =~ s|(<£[^>]*>[^>]*</£[^>]*>)|<label-g>$1</label-g>|gio;
	$e =~ s|£||g;
    }
    if ($dict =~ m|ald8e|i)
    {
	$e =~ s|(<cf[ >])|£$1|gi;
	$e =~ s|(<cf[^£]*</cf>) *(<x-g[^>]*>)|$2$1|gi;
	$e =~ s|£||g;
    }
    $e =~ s|</label-g>(<cm[^<]*</cm>)<label-g>|$1|gio;
    $e =~ s|</label-g><label-g>||gio;
    $e =~ s|<label-g>|£<label-g>|g;
    if ($e =~ s|(<label-g[^>]*>[^£]*</label-g>)(<x-g[^>]*>)|$2$1|g)
    {
	$label = $1;
	printf(log_fp "%s\t%s\n", $label, $h); 
    }
    $e =~ s|</x-g-tmp> *<x-g-tmp[^>]*>||gi;
    $e = restructure::tag_rename($e, "x-g-tmp", "x-g");
    $e =~ s|£||g;
    return $e;
}

sub add_xt_to_xr_g
{
    my($e) = @_;
    my($bit, $res, $xt);
    my(@BITS);
    unless ($e =~ m|<xr-g|i)
    {
	return $e;
    }
    $e =~ s|(<xr-g[ >].*?</xr-g>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    unless ($bit =~ m|<xr-g[^>]* xt=\"[a-z]|i)
	    {
		if ($bit =~ m| (xt=\"[^\"]*\")|i)
		{
		    $xt = $1;
		    $bit =~ s|(<xr-g)([ />])|$1 $xt $2|i;
		}
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_posg
{
    my($e, $dict) = @_;
    $e =~ s|(<p [^>]*/>)|<pos-g>$1</pos-g>|gio;
    if ($dict =~ m|oad|i)
    {
	while ($e =~ s|(</pos-g>)(<pt [^>]*/>)|$2$1|gi)
	{}
	while ($e =~ s|(</pos-g>)(<gr [^>]*/>)|$2$1|gi)
	{}
    }
    $e =~ s|</pos-g><pos-g>||gio;
    return $e;
}

sub add_xt_to_xrs
{
    my($e) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<xr-g[ >].*?</xr-g>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    $bit = &add_xt_to_xr($bit);
	}
	$res .= $bit;
    }
    return $res;
}


sub add_xt_to_xr
{
    my($e) = @_;
    my($bit, $res, $xt);
    my(@BITS);
    if ($e =~ m|<xr-g [^>]*xt=\"(.*?)\"|i)
    {
	$xt = $1;
    }
    $e =~ s|<xr>|<xr >|gi;
    $e =~ s|(<xr .*?>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    unless ($bit =~ m| xt=|)
	    {
		$bit =~ s|<xr |<xr xt=\"$xt\" |i;
	    }
	}
	$res .= $bit;
    }
    return $res;    
}

sub do_ndvs
{
    my ($e) = @_;
    my($res, $type, $bit, $xs, $xp);
    my(@BITS);    
    $e = &add_xt_to_xrs($e);
    if ($e =~ m|<xr-gs|)
    {
	$e = &convert_to_ndvs($e);
    }
    $e =~ s|<xr-g[^>]*xt=\"[^\"]*ndv\"[^>]*>(.*?)</xr-g>|$1|gio;
    $e =~ s|(<xr[ >].*?</xr>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS)
    {
	if ($bit =~ s|&fk;||gio)
	{
	    $type = "";
	    if ($bit =~ m|xt=\"(.*?)\"|io)
	    {
		$type = $1;
		$type =~ tr|A-Z|a-z|;
	    }
	    if ($type =~ m|ndv|io)
	    {
		$xs = "";
		$xhm = restructure::get_tag_contents($bit, "xhm"); 
		unless ($xhm =~ m|^ *$|)
		{
		    $xhm = sprintf("<sup >%s</sup>", $xhm); 
		}
		if ($bit =~ s|(<xp.*?</xp>)||i)
		{
		    $xp = $1;
		}
		if ($bit =~ s|(<xs.*?</xs>)||i)
		{
		    $xs = $1;
		}
		$bit =~ s|<[^>]*>||gio;
		$bit = sprintf("<ndv>%s%s%s</ndv>", $bit, $xhm, $xs);
	    }
	}
	$res .= $bit;
    }
    $res =~ s|([^> \(])(<ndv)|$1 $2|gio;
    $res =~ s|(</ndv>)([^<\.\!\'\?,\)\:; ])|$1 $2|gio;
    $res =~ s|(</ndv>)(<[^/z])|$1 $2|gio;
    $res = &lose_xh_from_ndvs($res);
    return $res;
}

sub lose_xh_from_ndvs
{
    my($e) = @_;
    my($res, $eid);	
    my($bit);
    my(@BITS);
    $e =~ s|(<ndv[ >].*?</ndv>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit = restructure::tag_rename($bit, "xhm", "sup");
	    $bit =~ s|</?x[hr].*?>||gi;
	}
	$res .= $bit;
    }    
    return $res;
}

sub convert_to_ndvs
{
    my($e) = @_;
    my($res, $eid);	
    my($bit);
    my(@BITS);
    $e =~ s|(<xr-gs[ >].*?</xr-gs>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $type = restructure::get_tag_attval($bit, "xr-gs", "xt"); 
	    if ($type =~ m|ndv|i)
	    {
		my $href = restructure::get_tag_attval($bit, "xr-g", "href"); 
		$bit =~ s|</?xr-gs[^>]*>||gi;
		$bit =~ s|<xr-g[^>]*>|<ndv>|gi;
		$bit =~ s|</xr-g[^>]*>|</ndv>|gi;
		unless ($bit =~ m|<ndv|i)
		{
		    $bit =~ s|^.*?<xh [^>]*>|<ndv>|i;
		    $bit =~ s|</xh>.*|</ndv>|i;
		}
		unless ($href =~ m|^ *$|)
		{
		    $bit = restructure::set_tag_attval($bit, "ndv", "href", $href); 
		}
	    }
	}
	$res .= $bit;
    }    
    return $res;
}


sub expand_tildes
{
    my($e) = @_;
    $e = &expand_group($e, "dr-g", "dr");
    $e = &expand_group($e, "entry", "h");
    return $e;
}


sub expand_group
{
    my($e, $group, $tag) = @_;
    my($bit, $res, $bold);
    my(@BITS);
    $e =~ s|(<$group[ >].*?</$group>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ m|(<$tag[ >].*?</$tag>)|i)
	    {
		$bold = $1;
		$bold =~ s|&[swp];||gi;
		$bold =~ s|<.*?>||g;
		$bold =~ s|&\#x2122;||gio; # &tm;
		$bold =~ s|&\#x2022;||gio; # &tm;
		$bold =~ s|&\#x2027;||gio; # &w;
		$bold =~ s|&\#x00B7;||gio; # &w;
		$bold =~ s|&#x02C[C8];||gio;  # &[ps]
#		$bold = sprintf("<exp >%s</exp>", $bold);
		$bit = &replace($bit, "li", $bold);
		$bit = &replace($bit, "x", $bold);
		$bit = &replace($bit, "cf", $bold);
		$bit = &replace($bit, "cfe", $bold);
		$bit = &replace($bit, "coll", $bold);
		$bit = &replace($bit, "v_alt", $bold);
		$bit = &replace($bit, "v", $bold);
	    }
	}
	$res .= $bit;
    }    
    return $res;
}

sub replace
{
    my($e, $tag, $bold) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<$tag[ >].*?</$tag>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s|(<exp [^>]*>) *</exp>|$1$bold</exp>|g;
	    $bit =~ s|~|<exp >$bold</exp>|g;
	}
	$res .= $bit;
    }
    return $res;
}

sub eb_in_if
{
    my($e) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<if[ >].*?</if>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    $bit =~ s| *<eb([^>]*>[^>]*)</eb> *|<co$1</co>|gi;
	}
	$res .= $bit;
    }
    return $res;
}

sub eb_in_x
{
    my($e) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<x[ >].*?</x>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    $bit =~ s| *<eb([^>]*>[^>]*)</eb> *|<eb_in_x$1</eb_in_x>|gi;
	}
	$res .= $bit;
    }
    return $res;
}

sub eb_in_d
{
    my($e) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<d[ >].*?</d>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    $bit =~ s| *<eb([^>]*>[^>]*)</eb> *|<eb_in_d$1</eb_in_d>|gi;
	}
	$res .= $bit;
    }
    return $res;
}


sub do_prac_pron
{
    my($e) = @_;
    my($bit, $res, $prac_pron);
    my(@BITS);
    undef %FUSED;
    $prac_pron = "";
    $h = &get_hdwd($e);
    $cp = $e;
    $e =~ s|(<h[ >].*?</h>)|&split;&fkbold;$1&split;|gio;
    $e =~ s|(<dr[ >].*?</dr>)|&split;&fkbold;$1&split;|gio;
    $e =~ s|(<zd[ >].*?</zd>)|&split;&fkbold;$1&split;|gio;
    $e =~ s|(<bf[ >].*?</bf>)|&split;&fkbold;$1&split;|gio;
    $e =~ s|(<if[ >].*?</if>)|&split;&fkbold;$1&split;|gio;
    $e =~ s|(<vp[ >].*?</vp>)|&split;&fkbold;$1&split;|gio;
    $e =~ s|(<v[ >].*?</v>)|&split;&fkbold;$1&split;|gio;
#    $e =~ s|<vps-g.*?</vps-g>||gi;
#    $e =~ s|<vs-g.*?</vs-g>||gi;
    $e =~ s|(<i-g[ >].*?</i-g>)|&split;&fk;$1&split;|gio;
    $e =~ s|(<ei-g[ >].*?</ei-g>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    $bold = "";
    $boldtype = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fkbold;||gio){
	    $bold = $bit;
	    $boldtype = $bit;
	    $boldtype =~ s|[ >].*$||;
	    $boldtype =~ s|</?||;
	    $boldtype =~ tr|A-Z|a-z|;
	    $bold =~ s|<.*?>||g;
	}
	if ($bit =~ s|&fk;||gio){
	    if ($bit =~ m| wd=\"(.*?)\"|)
	    {
		$wd = $1;
		$f = &get_audio_file($bit);		
		if (&wd_comp($bold, $wd))
		{
		    $wdf = sprintf("%s\t%s", $bold, $f);
		    unless ($FUSED{$wdf})
		    {
			$prac_pron = sprintf("%s<pron-g><wd type=\"%s\">%s</wd>%s</pron-g>", $prac_pron, $boldtype, $bold, $bit);
		    }
		    $FUSED{$wdf} = 1;
		}
		elsif (&wd_comp($h, $wd))
		{
		    $wdf = sprintf("%s\t%s", $h, $f);
		    unless ($FUSED{$wdf})
		    {
			$prac_pron = sprintf("%s<pron-g><wd type=\"%s\">%s</wd>%s</pron-g>", $prac_pron, $boldtype, $h, $bit);
		    }
		    $FUSED{$wdf} = 1;
		}
	    }
	}
    }    
    unless ($prac_pron =~ m|^ *$|)
    {
	$prac_pron = sprintf("<pracpron>%s</pracpron>", $prac_pron);
    }
    $res = $cp;
    $res =~ s|(</entry>)|$prac_pron$1|i;
    return $res;
}

sub get_spoken_examples
{
    my($e) = @_;
    my($bit, $res);
    my(@BITS3);
    $e =~ s|(<x-g[ >].*?</x-g>)|&split;&fk;$1&split;|gio;
    @BITS3 = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS3){
	if ($bit =~ s|&fk;||gio){
	    if (($bit =~ m|<e?i-g[ >]|i) || ($bit =~ m|<audio|i))
	    {
		$res .= $bit;
	    }
	}
    }
    unless ($res =~ m|^ *$|)
    {
	$res = sprintf("<spx-g>%s</spx-g>", $res);
    }
    return $res;
}

sub put_audio_on_ig_tags
{
    my($e) = @_;
    my($bit, $res);
    my(@BITS);
    my($wd);
    $e =~ s|(<e?i-g[ >].*?</e?i-g>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    if ($bit =~ m| (wd=\".*?\")|i)
	    {
		$wd = $1;
		unless ($bit =~ m|<i-g[^>]* wd=|)
		{
		    $bit =~ s|>| $wd>|;
		}
	    }
	    $bit = &get_first_soundfiles($bit);
	    $bit = &delete_dummies($bit);
	    if ($bit =~ m|<i-g[^>]*> *</i-g>|)
	    {
		$bit =~ s|(<i-g[^>]*)>|$1 slashes=\"n\">|i;
	    }
	}
	$res .= $bit;
    }   
    return $res;
}

sub get_first_soundfiles
{
    my($e) = @_;
    my($gb, $us);
    if ($e =~ m|<phon-|i)
    {
	if ($e =~ m|^.*?<phon-gb [^>]*file=(\".*?\")|i)
	{
	    $gb = $1;
	    $e =~ s|(<e?i-g[^>]*)>|$1 gbf=$gb>|i;
	}
	if ($e =~ m|^.*?<phon-us [^>]*file=(\".*?\")|i)
	{
	    $us = $1;
	    $e =~ s|(<e?i-g[^>]*)>|$1 usf=$us>|i;
	}
    }
    else
    {
	if ($e =~ m|^.*?<i [^>]*file=(\".*?\")|i)
	{
	    $gb = $1;
	    $e =~ s|(<e?i-g[^>]*)>|$1 gbf=$gb>|i;
	}
	if ($e =~ m|^.*?<y [^>]*file=(\".*?\")|i)
	{
	    $us = $1;
	    $e =~ s|(<e?i-g[^>]*)>|$1 usf=$us>|i;
	}
    }
    return($e);
}


sub delete_dummies
{
    my($e) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<phon-gb[ >].*?</phon-gb>)|&split;&fk;$1&split;|gio;
    $e =~ s|(<phon-us[ >].*?</phon-us>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|display=\"n|i)
	    {
		$bit = "";
	    }
	    if ($bit =~ m|<phon[^>]*> *</phon|i)
	    {
		$bit = "";
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub get_audio_file
{
    my($e) = @_;
    if ($bit =~ m| gbf=\"(.*?)\"|i)
    {
	$f = $1;
    }
    elsif ($bit =~ m| file=\"(.*?)\"|i)
    {
	$f = $1;
    }
    $f =~ s|^.*/||;
    $f =~ s|\..*||;
    $f =~ tr|A-Z|a-z|;
    return $f;
}

sub wd_comp
{
    my($wd1, $wd2) = @_;
    $wd1 =~ tr|A-Z|a-z|;
    $wd1 =~ s|,.*||;
    $wd2 =~ s|,.*||;
    $wd1 =~ s|&.*?;||g;
    $wd1 =~ s|[^a-z0-9]||g;
    $wd2 =~ tr|A-Z|a-z|;
    $wd2 =~ s|&.*?;||g;
    $wd2 =~ s|[^a-z0-9]||g;
    if ($wd1 eq $wd2)
    {
	return 1;
    }
    return 0;
}

sub add_levels_info
{
    my($e, $toptag) = @_;
    my(@BITS);
    my($res);
    $e =~ s| level=\"[^\"]*\"||gi;
    if ($toptag =~ m|^ *$|)
    {
	$toptag = "entry";
    }
    unless ($e =~ m|<$toptag|i)
    {
	return $e;
    }
    $e =~ s|(<[a-z/].*?>)|&split;&fk;$1&split;|gio;
    my $level = 0;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|</|)
	    {
		$level--;
	    }
	    else
	    {
		$bit =~ s|(<[^> /]*)|$1 level=\"$level\"|;
		unless ($bit =~ m|/ *>|)
		{
		    $level++;
		}
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_parent_info
{
    # requires add_levels_first
    my($e) = @_;
    my($bit, $res);
    my(@BITS);
    my %TNAMES;
    $e =~ s|(<[a-z][^>]*>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<([^ /]*).* level=\"([^\"]*)\"|i)
	    {
		$tname = $1;
		$level = $2;
		$plevel = $level - 1;
		$parent = $TNAME{$plevel};
		$TNAME{$level} = $tname;
		unless ($parent =~ m|^ *$|)
		{
		    $bit =~ s| | parent=\"$parent\" |i;
		}
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub add_zp_to_tag_following_unbox
{
    my($e) = @_;
    unless ($e =~ m|<unbox|i)
    {
	return $e;
    }
    my($bit, $res);
    my(@BITS);
    $e =~ s|(</unbox>)|&split;&fk;$1|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    $bit =~ s|(<[a-z][^ />]*)([ />])|$1 zp=\"y\"$2|i;
	}
	$res .= $bit;
    }
    return $res;
}

##################################################

sub rename_tag_in_tag
{
    my($e, $container, $oldtag, $newtag, $tofix, $psg) = @_;

    my($bit, $res);
    my(@BITS);
    $e =~ s|(<$container[ >].*?</$container>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    $bit =~ s|<$oldtag([ >])|<$newtag$1|gi;
	    $bit =~ s|</$oldtag([ >])|</$newtag$1|gi;
	    unless ($tofix =~ m|^ *$|)
	    {
#		$tofix = sprintf(" tofix=\"$tofix\"");
		$bit = restructure::set_tag_attval($bit, $newtag, "tofix", $tofix); 
		$bit = &set_psg_attval($bit, $newtag, $psg);
	    }
	    
	}
	$res .= $bit;
    }
    return $res;
}

sub lose_unicode_ents
{
    my($e) = @_;
    $e =~ s|&\#x2027;||gio; # &w;
    $e =~ s|&#x02C[C8];||gio;  # &[ps]
    $e =~ s|&#x2009;| |gio;
    $e =~ s|&#x2013;| |gio;
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

sub get_hdwd
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
#	$h = &change_ae_ents($h);
	$h = &ents2letter($h);
	return $h;
    }
    else
    {
	return $file_hdwd;
    }
}


sub hex2letter
{
    my($e) = @_;
    unless ($e =~ m|&\#x|)
    {
	return $e;
    }
    $e =~ s|&\#x00C6;|AE|go;
    $e =~ s|&\#x00C0;|A|go;
    $e =~ s|&\#x00C1;|A|go;
    $e =~ s|&\#x00C2;|A|go;
    $e =~ s|&\#x00C3;|A|go;
    $e =~ s|&\#x00C4;|A|go;
    $e =~ s|&\#x00C5;|A|go;
    $e =~ s|&\#x00C7;|C|go;
    $e =~ s|&\#x00C8;|E|go;
    $e =~ s|&\#x00C9;|E|go;
    $e =~ s|&\#x00CA;|E|go;
    $e =~ s|&\#x00CB;|E|go;
    $e =~ s|&\#x00CC;|I|go;
    $e =~ s|&\#x00CD;|I|go;
    $e =~ s|&\#x00CE;|I|go;
    $e =~ s|&\#x00CF;|I|go;
    $e =~ s|&\#x00D1;|N|go;
    $e =~ s|&\#x00D2;|O|go;
    $e =~ s|&\#x00D3;|O|go;
    $e =~ s|&\#x00D4;|O|go;
    $e =~ s|&\#x00D5;|O|go;
    $e =~ s|&\#x00D6;|O|go;
    $e =~ s|&\#x00D8;|O|go;
    $e =~ s|&\#x00D9;|U|go;
    $e =~ s|&\#x00DA;|U|go;
    $e =~ s|&\#x00DB;|U|go;
    $e =~ s|&\#x00DC;|U|go;
    $e =~ s|&\#x00DD;|Y|go;
    $e =~ s|&\#x00DF;|sz|go;
    $e =~ s|&\#x00E0;|a|go;
    $e =~ s|&\#x00E1;|a|go;
    $e =~ s|&\#x00E2;|a|go;
    $e =~ s|&\#x00E3;|a|go;
    $e =~ s|&\#x00E4;|a|go;
    $e =~ s|&\#x00E5;|a|go;
    $e =~ s|&\#x00E6;|ae|go;
    $e =~ s|&\#x00E7;|c|go;
    $e =~ s|&\#x00E8;|e|go;
    $e =~ s|&\#x00E9;|e|go;
    $e =~ s|&\#x00EA;|e|go;
    $e =~ s|&\#x00EB;|e|go;
    $e =~ s|&\#x00EC;|i|go;
    $e =~ s|&\#x00ED;|i|go;
    $e =~ s|&\#x00EE;|i|go;
    $e =~ s|&\#x00EF;|i|go;
    $e =~ s|&\#x00F1;|n|go;
    $e =~ s|&\#x00F2;|o|go;
    $e =~ s|&\#x00F3;|o|go;
    $e =~ s|&\#x00F4;|o|go;
    $e =~ s|&\#x00F5;|o|go;
    $e =~ s|&\#x00F6;|o|go;
    $e =~ s|&\#x00F9;|u|go;
    $e =~ s|&\#x00FA;|u|go;
    $e =~ s|&\#x00FB;|u|go;
    $e =~ s|&\#x00FC;|u|go;
    $e =~ s|&\#x00FD;|y|go;
    $e =~ s|&\#x00FF;|y|go;
    $e =~ s|&\#x0100;|A|go;
    $e =~ s|&\#x0101;|a|go;
    $e =~ s|&\#x0102;|A|go;
    $e =~ s|&\#x0103;|a|go;
    $e =~ s|&\#x0104;|A|go;
    $e =~ s|&\#x0105;|a|go;
    $e =~ s|&\#x0106;|C|go;
    $e =~ s|&\#x0107;|c|go;
    $e =~ s|&\#x0108;|C|go;
    $e =~ s|&\#x0109;|c|go;
    $e =~ s|&\#x010A;|C|go;
    $e =~ s|&\#x010B;|c|go;
    $e =~ s|&\#x010C;|C|go;
    $e =~ s|&\#x010D;|c|go;
    $e =~ s|&\#x010E;|D|go;
    $e =~ s|&\#x010F;|d|go;
    $e =~ s|&\#x0110;|D|go;
    $e =~ s|&\#x0111;|d|go;
    $e =~ s|&\#x0112;|E|go;
    $e =~ s|&\#x0113;|e|go;
    $e =~ s|&\#x0116;|E|go;
    $e =~ s|&\#x0117;|e|go;
    $e =~ s|&\#x0118;|E|go;
    $e =~ s|&\#x0119;|e|go;
    $e =~ s|&\#x011A;|E|go;
    $e =~ s|&\#x011B;|e|go;
    $e =~ s|&\#x011C;|G|go;
    $e =~ s|&\#x011D;|g|go;
    $e =~ s|&\#x011E;|G|go;
    $e =~ s|&\#x011F;|g|go;
    $e =~ s|&\#x0120;|G|go;
    $e =~ s|&\#x0121;|g|go;
    $e =~ s|&\#x0122;|G|go;
    $e =~ s|&\#x0124;|H|go;
    $e =~ s|&\#x0125;|h|go;
    $e =~ s|&\#x0126;|H|go;
    $e =~ s|&\#x0127;|h|go;
    $e =~ s|&\#x0128;|I|go;
    $e =~ s|&\#x0129;|i|go;
    $e =~ s|&\#x012A;|I|go;
    $e =~ s|&\#x012B;|i|go;
    $e =~ s|&\#x012E;|I|go;
    $e =~ s|&\#x012F;|i|go;
    $e =~ s|&\#x0130;|I|go;
    $e =~ s|&\#x0131;|i|go;
    $e =~ s|&\#x0132;|IJ|go;
    $e =~ s|&\#x0133;|ij|go;
    $e =~ s|&\#x0134;|J|go;
    $e =~ s|&\#x0135;|j|go;
    $e =~ s|&\#x0136;|K|go;
    $e =~ s|&\#x0137;|k|go;
    $e =~ s|&\#x0139;|L|go;
    $e =~ s|&\#x013A;|l|go;
    $e =~ s|&\#x013B;|L|go;
    $e =~ s|&\#x013C;|l|go;
    $e =~ s|&\#x013D;|L|go;
    $e =~ s|&\#x013E;|l|go;
    $e =~ s|&\#x013F;|L|go;
    $e =~ s|&\#x0140;|l|go;
    $e =~ s|&\#x0141;|L|go;
    $e =~ s|&\#x0142;|l|go;
    $e =~ s|&\#x0143;|N|go;
    $e =~ s|&\#x0144;|n|go;
    $e =~ s|&\#x0145;|N|go;
    $e =~ s|&\#x0146;|n|go;
    $e =~ s|&\#x0147;|N|go;
    $e =~ s|&\#x0148;|n|go;
    $e =~ s|&\#x014C;|O|go;
    $e =~ s|&\#x014D;|o|go;
    $e =~ s|&\#x0152;|OE|go;
    $e =~ s|&\#x0153;|oe|go;
    $e =~ s|&\#x0154;|R|go;
    $e =~ s|&\#x0155;|r|go;
    $e =~ s|&\#x0156;|R|go;
    $e =~ s|&\#x0157;|r|go;
    $e =~ s|&\#x0158;|R|go;
    $e =~ s|&\#x0159;|r|go;
    $e =~ s|&\#x015A;|S|go;
    $e =~ s|&\#x015B;|s|go;
    $e =~ s|&\#x015C;|S|go;
    $e =~ s|&\#x015D;|s|go;
    $e =~ s|&\#x015E;|S|go;
    $e =~ s|&\#x015F;|s|go;
    $e =~ s|&\#x0160;|S|go;
    $e =~ s|&\#x0161;|s|go;
    $e =~ s|&\#x0162;|T|go;
    $e =~ s|&\#x0163;|t|go;
    $e =~ s|&\#x0164;|T|go;
    $e =~ s|&\#x0165;|t|go;
    $e =~ s|&\#x0166;|T|go;
    $e =~ s|&\#x0167;|t|go;
    $e =~ s|&\#x0168;|U|go;
    $e =~ s|&\#x0169;|u|go;
    $e =~ s|&\#x016A;|U|go;
    $e =~ s|&\#x016B;|u|go;
    $e =~ s|&\#x016C;|U|go;
    $e =~ s|&\#x016D;|u|go;
    $e =~ s|&\#x016E;|U|go;
    $e =~ s|&\#x016F;|u|go;
    $e =~ s|&\#x0172;|U|go;
    $e =~ s|&\#x0173;|u|go;
    $e =~ s|&\#x0174;|W|go;
    $e =~ s|&\#x0175;|w|go;
    $e =~ s|&\#x0176;|Y|go;
    $e =~ s|&\#x0177;|y|go;
    $e =~ s|&\#x0178;|Y|go;
    $e =~ s|&\#x0179;|Z|go;
    $e =~ s|&\#x017A;|z|go;
    $e =~ s|&\#x017B;|Z|go;
    $e =~ s|&\#x017C;|z|go;
    $e =~ s|&\#x017D;|Z|go;
    $e =~ s|&\#x017E;|z|go;
    $e =~ s|&\#x01F5;|g|go;
    $e =~ s|&\#xFB00;|ff|go;
    $e =~ s|&\#xFB01;|fi|go;
    $e =~ s|&\#xFB02;|fl|go;
    $e =~ s|&\#xFB03;|ffi|go;
    $e =~ s|&\#xFB04;|ffl|go;
    $e =~ s|&\#x[^ ;]*;||go;
    return $e;
}

sub ents2letter
{
    my($e) = @_;
    $e = &lose_unicode_ents($e);
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


sub renumber_ngs
{
    my($e) = @_;
    my($res, $bit, $expected, $num);
    $res = "";
    my(@BITS);
    $e =~ s/(<bf-g [^>]*>)/&split;&fk;$1&split;/gi;
    $e =~ s/(<dr-g [^>]*>)/&split;&fk;$1&split;/gi;
    $e =~ s/(<h-g [^>]*>)/&split;&fk;$1&split;/gi;
    $e =~ s/(<id-g [^>]*>)/&split;&fk;$1&split;/gi;
    $e =~ s/(<p-g [^>]*>)/&split;&fk;$1&split;/gi;
    $e =~ s/(<pv-g [^>]*>)/&split;&fk;$1&split;/gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    $bit = &renumber_group($bit);
	}	
	$res .= $bit;
    }
    return $res;
}

sub renumber_group
{
    my($e) = @_;
    my($bit, $res, $expected);
    my(@BITS);
    $expected = 0;
    $e =~ s|<n-g>|<n-g >|gi;
    $e =~ s|(<n-g[ >].*?</n-g>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gio){
	    $expected++;
	    unless ($bit =~ m|<n-g[^>]* n=\"|)
	    {
		$bit =~ s|<n-g |<n-g n=\"99\" |i;
	    }
	    if ($bit =~  m|<n-g[^>]* n=\"([^\"]*)\"|)
	    {
		$num = $1;
		unless ($num == $expected)
		{
		    $bit =~ s| n=\".*?\"| n=\"$expected\"|i;
		}
	    }
	}
	$res .= $bit;
    }
    unless ($res =~ m|<n-g.*<n-g|i)
    {
	$res =~ s|<n-g |<sense-g |i;
	$res =~ s|</n-g>|</sense-g>|i;
    }
    return $res;
}

sub catname
{
    my($e, $tagname, $oldatt, $newatt) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<$tagname .*?>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s| $oldatt=| $newatt=|gi;
	}
	$res .= $bit;
    }
    return $res;
}

# $e = restructure::del_attrib($e, "TAGNAME", "ATTNAME", "TOFIX", "PSGCOMMENT");
sub del_attrib
{
    my($e, $tagname, $attname, $tofix, $psg) = @_;
#    $e = &catval($e, "unbox", "unbox", "coll", "colloc");    
    my($bit, $res);
    my(@BITS);
    unless ($tofix =~ m|^ *$|)
    {
	$tofix = sprintf(" tofix=\"$tofix\"");
    }
    $e =~ s|(<$tagname .*?>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ s| $attname=\"[^\"]*\"| $tofix|gi)
	    {
		$bit = &set_psg_attval($bit, $tagname, $psg);
	    }
	}
	$res .= $bit;
    }
    return $res;
}



sub catval
{
    my($e, $tagname, $attname, $oldval, $newval, $tofix, $psg) = @_;
#    $e = &catval($e, "unbox", "unbox", "coll", "colloc");    
    my($bit, $res);
    my(@BITS);
    unless ($tofix =~ m|^ *$|)
    {
	$tofix = sprintf(" tofix=\"$tofix\"");
    }
    $e =~ s|(<$tagname .*?>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ s| $attname=\"$oldval\"| $attname=\"$newval\"$tofix|gi)
	    {
		$bit = &set_psg_attval($bit, $tagname, $psg);
	    }
	}
	$res .= $bit;
    }
    return $res;
}


sub setqattrib
{
    my($e, $tagname, $attname, $oldval, $qval, $newval, $tofix, $psg) = @_;
#    $e = &catval($e, "unbox", "unbox", "coll", "colloc");    
    my($bit, $res);
    my(@BITS);
    unless ($tofix =~ m|^ *$|)
    {
	$tofix = sprintf(" tofix=\"$tofix\"");
    }
    $e =~ s|(<$tagname .*?>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ s| $attname=\"$oldval\"| q=\"$qval\" $attname=\"$newval\"$tofix|gi)
	    {
		$bit = &set_psg_attval($bit, $tagname, $psg);
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub loseattrib
{
    my($e, $tagname, $attname) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<$tagname .*?>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s| $attname=\"[^\"]*\"||gi;
	}
	$res .= $bit;
    }
    return $res;
}

sub tagrename
{
    my($e, $oldtag, $newtag) = @_;
    $e =~ s|<$oldtag */>|<$newtag />|gi;
    $e =~ s|<$oldtag([ />])|<$newtag$1|gi;
    $e =~ s|</$oldtag *>|</$newtag>|gi;
    return $e;
}

sub tag_rename
{
    my($e, $oldtag, $newtag, $tofix, $psg) = @_;
    unless ($tofix =~ m|^ *$|)
    {
	$tofix = sprintf(" tofix=\"%s\"", $tofix);
    }
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<$oldtag[ >].*?</$oldtag>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s|<$oldtag([ />])|<$newtag$tofix$1|gi;
	    if ($bit =~ s|</$oldtag *>|</$newtag>|gi)
	    {
		$bit = &set_psg_attval($bit, $newtag, $psg);
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub tag_delete
{
    my($e, $tagname) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<$tagname */>)||gi;
    $e =~ s|(<$tagname [^>]*/>)||gi;
    $e =~ s|(<$tagname[ >].*?</$tagname>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit = "";
	}
	$res .= $bit;
    }    
    return $res;
}


sub tag_delete_empty
{
    my($e, $tagname) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|<$tagname>|<$tagname >|gi;
    $e =~ s|(<$tagname [^>]*> *</$tagname>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit = "";
	}
	$res .= $bit;
    }    
    return $res;
}


sub lose_tag
{
    my($e, $tagname) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<$tagname)([^a-z0-9_\- ])|$1 $2|gi;
    $e =~ s|(<$tagname [^>]*>)||gi;
    $e =~ s|(</$tagname>)||gi;
    return $e;
}

sub attrib_rename
{
    my($e, $tag, $oldatt, $newatt) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<$tag .*?>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s| $oldatt=| $newatt=|i;
	}
	$res .= $bit;
    }
    return $res;
}


sub get_attval
{
    my($e, $att) = @_;
    my($bit);
    if ($e =~ m| $att=\"(.*?)\"|i)
    {
	return $1;
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

sub get_tag_contents
{
    my($e, $tag) = @_;
    my($res);
    $e =~ s|<$tag>|<$tag >|gi;
    if ($e =~ m|<$tag [^>]*>(.*?)</$tag>|i)
    {
	$res = $1;
	return $res;
    }
    return "";
}

sub get_tag
{
    my($e, $tag) = @_;
    my($res);
    $e =~ s|<$tag>|<$tag >|gi;
    if ($e =~ m|(<$tag [^>]*>.*?</$tag>)|i)
    {
	$res = $1;
	return $res;
    }
    return "";
}


sub get_tagname
{
    my($e) = @_;
    my($res);
    $e =~ s|[ >].*$||i;
    $e =~ s|<||i;
    return($e);
}


sub set_psg_attval
{
    my($e, $tagname, $newval) = @_;
#    $bit = &set_att_val($bit, "xr-g", "href", "TARGET", "colloc");    
    my($bit, $res, $oldpsg, $attval);
    if ($newval =~ m|^ *$|)
    {
	return $e;
    }
    $oldpsg = restructure::get_tag_attval($e, $tagname, "psg"); 
    $attval = sprintf("%s{%s}", $oldpsg, $newval);
    $e = restructure::set_tag_attval($e, $tagname, "psg", $attval); 
    return $e;
}


sub set_tag_attval
{
    my($e, $tagname, $attname, $newval, $tofix, $psg) = @_;
#    $bit = &set_att_val($bit, "xr-g", "href", "TARGET", "colloc");    
    my($bit, $res);
    my(@BITS);
    $e =~ s|<$tagname>|<$tagname >|gi;
    unless ($tofix =~ m|^ *$|)
    {
	$tofix = sprintf(" tofix=\"$tofix\"");
    }
    $e =~ s|(<$tagname .*?>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s| $attname=\"[^\"]*\"| |;
	    if ($bit =~ s| | $attname=\"$newval\"$tofix |i)
	    {
		$bit = &set_psg_attval($bit, $tagname, $psg);
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub set_first_tag_attval
{
    my($e, $tagname, $attname, $newval, $tofix, $psg) = @_;
#    $bit = &set_att_val($bit, "xr-g", "href", "TARGET", "colloc");    
    my($bit, $res);
    my(@BITS);
    # Rename the first occurrence of the tag to a temporary name, then change it's attrib value and change it back at the end.
    $e =~ s|<$tagname>|<$tagname >|gi;
    unless ($tofix =~ m|^ *$|)
    {
	$tofix = sprintf(" tofix=\"$tofix\"");
    }
    my $tmp_tagname = sprintf("%s_fktmp", $tagname); 
    $e =~ s|<$tagname |<$tmp_tagname |;
    $e =~ s|</$tagname>|</$tmp_tagname>|;
    $e =~ s|(<$tmp_tagname .*?>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s| $attname=\"[^\"]*\"| |;
	    if ($bit =~ s| | $attname=\"$newval\"$tofix |i)
	    {
		$bit = &set_psg_attval($bit, $tmp_tagname, $psg);
	    }
	}
	$res .= $bit;
    }
    $res = restructure::tag_rename($res, $tmp_tagname, $tagname);
    return $res;
}

sub hide_comments
{
    my($e) = @_;
    my($bit, $res);
    my(@BITS);
    unless ($e =~ m|<!--(.*?)-->|gio)
    {
	return $e;
    }
    $e = &neutralise_comments($e);
    $e =~ s|(<[a-z])|&split;&fk;$1|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ m|<!--|)
	    {
		$comment = &get_comments($bit);
		$bit =~ s| *(/?>)| comments=\"$comment\" $1|;
	    }
	}
	$res .= $bit;
    }
    $res =~ s|<!--.*?-->||gio;
    return $res;
}


sub restore_comments
{
    my($e) = @_;
    my($res, $eid);	
    my($bit, $res);
    my(@BITS);
    return $e unless ($e =~ m| comments=|);
    $e =~ s|(<[^>]* comments=[^>]*>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s| comments=\"([^\"]*)\"(.*)$| \2<!-- \1 -->|;
	    $bit =~ s| +| |g;
	}
	$res .= $bit;
    }    
    while ($res =~ s|(<h [^>]*>)(<!--[^>]*-->)|\2\1|gi)
    {}
    while ($res =~ s|(<pos [^>]*>)(<!--[^>]*-->)|\2\1|gi)
    {}
    return $res;
}
sub neutralise_comments
{
    my($e) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<!--.*?-->)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s|<!-- *||;
	    $bit =~ s| *-->||;
	    $bit =~ s|&([^\#])|&amp;$1|g;
	    $bit =~ s|&amp;amp;|&amp;|gi;
	    $bit =~ s|<|&lt;|g;
	    $bit =~ s|>|&gt;|g;
	    $bit =~ s|[\'\"]||g;
	    $bit = sprintf("<!-- %s -->", $bit);
	}
	$res .= $bit;
    }
    return $res;
}

sub get_comments
{
    my($e) = @_;
    my($bit, $res);
    my(@BITS);
    $e =~ s|(<!--.*?-->)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s|<!-- *||;
	    $bit =~ s| *-->|&comsep;|;
	    while ($bit =~ s|\-\-+|\-|g)
	    {}
	    $res .= $bit;
	}
    }    
    $res =~ s|&comsep; *$||;
    $res =~ s|&comsep; *|; |;
    $res =~ s|~|&\#x0098;|gi;
    $res =~ s|  +| |g;
    return $res;
}

sub group_pgs
{
    my($e) = @_;
    my($res);	
    return $e unless ($e =~ m|<p-g[ >]|i);
    # change the logic so that if there are already <idm-gs tags then they are not repeated
    my(@BITS);
    $e =~ s|(<p-gs[ >].*?</p-gs>)|&split;&fk;$1&split;|gio;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gio){
	    if ($bit =~ m|<p-g[ >]|)
	    {
		$bit =~ s|(<p-g[ >].*?</p-g>)|<p-gs >$1</p-gs>|gio;
		$bit =~ s|</p-gs><p-gs [^>]*>||gio;
	    }
	}
	$res .= $bit;
    }
    return $res;
}


sub move_forward_out_of
{
    my($e, $prev, $tag) = @_;
    my($res, $eid);	
    $e =~ s|<$tag>|<$tag >|gi;
    $e =~ s|<$prev>|<$prev >|gi;
    $e =~ s|(<$prev )|£$1|gi;
    while ($e =~ m|</$prev> *</$tag>|i)
    {
	$e =~ s|(<$prev[^£]*</$prev>) *(</$tag>)|$2$1|gi;
    }
    $e =~ s|£||g;
    return $e;
}


sub move_forward_into
{
    my($e, $prev, $tag) = @_;
    my($res, $eid);	
    $e =~ s|<$tag>|<$tag >|gi;
    $e =~ s|(<$prev )|£$1|gi;
    while ($e =~ m|</$prev> *<$tag |i)
    {
	$e =~ s|(<$prev[^£]*</$prev>) *(<$tag [^>]*>)|$2$1|gi;
    }
    $e =~ s|£||g;
    return $e;
}

sub move_back_out_of
{
    my($e, $prev, $tag) = @_;
    my($res, $eid);	
    $e =~ s|<$tag>|<$tag >|gi;
    $e =~ s|<$prev>|<$prev >|gi;
    $e =~ s|(</$tag>)|$1£|gi;
    while ($e =~ m|<$prev [^>]*> *<$tag |i)
    {
	$e =~ s|(<$prev [^>]*>) *(<$tag [^£]*</$tag>)£*|$2$1|gi;
    }
    $e =~ s|£||g;
    return $e;
}

sub move_back_into
{
    my($e, $prev, $tag) = @_;
    my($res, $eid);	
    $e =~ s|<$tag>|<$tag >|gi;
    unless ($e =~ m|</$prev> *<$tag |)
    {
	return $e;
    }
    my($bit);
    my(@BITS);
    $e =~ s|(</$prev[ >])|&split;&fk;$1|gi;
    $e =~ s|(</$tag>)|$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    if ($bit =~ m|</$prev> *<$tag |i)
	    {
		$bit =~ s|(</$prev>) *(<$tag [^>]*>.*?</$tag>)|$2$1|gi;
	    }
	}
	$res .= $bit;
    }
    return $res;
}

sub move_back_into_old
{
    my($e, $prev, $tag) = @_;
    my($res, $eid);	
    $e =~ s|</$tag>|</$tag>£|gi;
    while ($e =~ m|</$prev> *<$tag[ >]|i)
    {
	$e =~ s|(</$prev>) *(<$tag[ >][^£]*</$tag>)£*|$2$1|gi;
    }
    $e =~ s|£||g;
    return $e;
}

sub add_missing_attvals
{
    my($e, $tagname, $attname, $genericval) = @_;
    my($res, $eid);	
    my($bit, $attval);
    my(@BITS);
    $e =~ s|(<$tagname[ >].*?</$tagname>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $val = restructure::get_tag_attval($bit, $tagname, $attname);    
	    if ($attval =~ m|^ *$|)
	    {
		$bit = restructure::set_tag_attval($bit, $tagname, $attname, $genericval); 
	    }
	}
	$res .= $bit;
    }    
    return $res;
}


sub extract_tag
{
    my($e, $tagname) = @_;
    my($res, $extract);	
    my($bit);
    my(@BITS);
    $e =~ s|(<$tagname [^>]*)/>|$1></$tagname>|gi;
    $e =~ s|(<$tagname[ >].*?</$tagname>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $extract .= $bit;
	    $bit = "";
	}
	$res .= $bit;
    }    
    return ($res, $extract);
}

sub expand_sb_sth
{
    my($e) = @_;
    my($res, $eid);	
    $e =~ s|£|&\#x00A3;|gio;
    $e = sprintf("£%s£", $e);
    my($bit);
    my(@BITS);
    $e =~ s|(<pron-gs[ >].*?</pron-gs>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	unless ($bit =~ s|&fk;||gi){
	    $bit =~ s|([^a-z])sb([^a-z])|$1somebody$2|gi;
	    $bit =~ s|([^a-z])sth([^a-z])|$1something$2|gi;
	}
	$res .= $bit;
    }    
    $res =~ s|£||g;
    return $res;
}

sub clean_commments
{
    my($e) = @_;
    my($res, $eid);	
    $e =~ s|<!\-\-(.*?)-->|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $bit =~ s|\-|_|gi;
	    $bit =~ s|<|&lt;|gi;
	    $bit =~ s|>|&gt;|gi;
	    $bit =~ s|[^a-z0-9&\#; ]+| |gi;
	    $bit = sprintf("<!--%s-->", $bit);
	}
	$res .= $bit;
    }
    return $res;
}

sub move_attrib
{
    my($e, $group, $fromtag, $totag, $attname) = @_;
    my($res, $eid);	
    my($bit, $attval, $currentval);
    my(@BITS);
    $e =~ s|(<$group[ >].*?</$group>)|&split;&fk;$1&split;|gi;
    @BITS = split(/&split;/, $e);
    $res = "";
    foreach $bit (@BITS){
	if ($bit =~ s|&fk;||gi){
	    $attval = restructure::get_tag_attval($bit, $fromtag, $attname); 
	    unless ($attval =~ m|^ *$|)
	    {
		$bit = restructure::del_attrib($bit, $fromtag, $attname);
		$currentval = restructure::get_tag_attval($bit, $totag, $attname); 
		if ($currentval =~ m|^ *$|)
		{
		    $bit = restructure::set_tag_attval($bit, $totag, $attname, $attval); 
		}
	    }
	}
	$res .= $bit;
    }    
    return $res;
}

sub reduce_prongs
{
    my($e) = @_;
    my($res, $eid);	
#<pron-gs wd="throw" source="ald8" psg="ald8_throw_prongs_1" eid="throw_prongs_1" e:id="u4cdebea65f7df6b4.-65ea7b53.14b82d3ebf8.-7097" e:target="297210" e:targetid="u4cdebea65f7df6b4.65a4283b.1430b4883ee.-a5f" e:targetproject="PHONETICS" e:targeteltid="u4cdebea65f7df6b4.65a4283b.1430b4883ee.-a5b">
    $e = restructure::del_attrib($e, "pron-gs", "e:target");
    $e = restructure::del_attrib($e, "pron-gs", "e:targetid");
    $e = restructure::del_attrib($e, "pron-gs", "e:targetproject");
    $e = restructure::del_attrib($e, "pron-gs", "e:targeteltid");
    $e = restructure::del_attrib($e, "phon", "e:inline");
    $e = restructure::del_attrib($e, "audio", "e:mediarecord");
    $e = restructure::del_attrib($e, "audio", "mediarecord");
    $e =~ s| xmlns:[a-z]*=\"urn:[^\" ]*\"| |g;
    $e =~ s| +| |g;
    return $e;
}



1;
