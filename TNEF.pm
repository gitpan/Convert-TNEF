# Convert::TNEF.pm
#
# Copyright (c) 1999 Douglas Wilson <dwilson@gtemail.net>. All rights reserved.
# This program is free software; you can redistribute it and/or
# modify it under the same terms as Perl itself.

# All attribute data will be just like a MIME::Body
# since its a class for storing stuff in either
# a file or a scalar.
require MIME::Body;

package Convert::TNEF::Scalar;
use vars qw(@ISA);
@ISA=qw(MIME::Body::Scalar);

package Convert::TNEF::File;
use vars qw(@ISA);
@ISA=qw(MIME::Body::File);

# Here's the main package (but NOT 'package main' ;^)
package Convert::TNEF;

use strict;
use vars qw(
 @ISA @EXPORT @EXPORT_OK $VERSION
 $TNEF_SIGNATURE
 $TNEF_PURE
 $LVL_MESSAGE
 $LVL_ATTACHMENT
 $errstr
 $g_file_cnt
 %dflts
 %atp
 %att
 %att_name
);

use Carp;
use IO::Wrap;
use File::Spec;

require AutoLoader;

@ISA = qw(Exporter AutoLoader);

$VERSION = '0.02';

# Preloaded methods go here.

# Set some TNEF constants. Everything turned
# out to be in little endian order, so I just added
# 'reverse' everywhere that I needed to
# instead of reversing the hex codes.
$TNEF_SIGNATURE = reverse pack('H*','223E9F78');
$TNEF_PURE      = reverse pack('H*','00010000');

$LVL_MESSAGE    = pack('H*','01');
$LVL_ATTACHMENT = pack('H*','02');

$atp{Triples}     = pack('H*','0000');
$atp{String}      = pack('H*','0001');
$atp{Text}        = pack('H*','0002');
$atp{Date}        = pack('H*','0003');
$atp{Short}       = pack('H*','0004');
$atp{Long}        = pack('H*','0005');
$atp{Byte}        = pack('H*','0006');
$atp{Word}        = pack('H*','0007');
$atp{Dword}       = pack('H*','0008');
$atp{Max}         = pack('H*','0009');

for (keys %atp) {
 $atp{$_} = reverse $atp{$_};
}

sub ATT {
 my ($att, $id) = @_;
 return reverse($id) . $att;
}

# The side comments are 'MAPI' equivalents
$att{Null}           = ATT( pack('H*','0000'), pack('H4','0000'));
$att{From}=ATT($atp{Triples}, pack('H*','8000')); # PR_ORIGINATOR_RETURN_ADDRESS
$att{Subject}        = ATT( $atp{String}, pack('H*','8004')); # PR_SUBJECT
$att{DateSent}    = ATT( $atp{Date}, pack('H*','8005')); # PR_CLIENT_SUBMIT_TIME
$att{DateRecd} = ATT( $atp{Date}, pack('H*','8006')); # PR_MESSAGE_DELIVERY_TIME
$att{MessageStatus}  = ATT( $atp{Byte}, pack('H*','8007')); # PR_MESSAGE_FLAGS
$att{MessageClass}   = ATT( $atp{Word}, pack('H*','8008')); # PR_MESSAGE_CLASS
$att{MessageID}      = ATT( $atp{String}, pack('H*','8009')); # PR_MESSAGE_ID
$att{ParentID}       = ATT( $atp{String}, pack('H*','800A')); # PR_PARENT_ID
$att{ConversationID}=ATT( $atp{String}, pack('H*','800B')); # PR_CONVERSATION_ID
$att{Body}           = ATT( $atp{Text}, pack('H*','800C')); # PR_BODY
$att{Priority}       = ATT( $atp{Short}, pack('H*','800D')); # PR_IMPORTANCE
$att{AttachData}     = ATT( $atp{Byte}, pack('H*','800F')); # PR_ATTACH_DATA_xxx
$att{AttachTitle}  = ATT( $atp{String}, pack('H*','8010')); # PR_ATTACH_FILENAME
$att{AttachMetaFile}= ATT( $atp{Byte}, pack('H*','8011')); # PR_ATTACH_RENDERING
$att{AttachCreateDate} = ATT( $atp{Date}, pack('H*','8012')); # PR_CREATION_TIME
$att{AttachModifyDate} = ATT( $atp{Date}, pack('H*','8013')); # PR_LAST_MODIFICATION_TIME
$att{DateModified}=ATT($atp{Date},pack('H*','8020'));# PR_LAST_MODIFICATION_TIME
$att{AttachTransportFilename} = ATT($atp{Byte}, pack('H*','9001')); #PR_ATTACH_TRANSPORT_NAME
$att{AttachRenddata}   = ATT( $atp{Byte}, pack('H*','9002'));
$att{MAPIProps}   = ATT( $atp{Byte}, pack('H*','9003'));
$att{RecipTable}  = ATT( $atp{Byte}, pack('H*','9004')); # PR_MESSAGE_RECIPIENTS
$att{Attachment}  = ATT( $atp{Byte}, pack('H*','9005'));
$att{TnefVersion} = ATT( $atp{Dword}, pack('H*','9006'));
$att{OemCodepage} = ATT( $atp{Byte}, pack('H*','9007'));
$att{OriginalMessageClass} = ATT( $atp{Word}, pack('H*','0006')); # PR_ORIG_MESSAGE_CLASS

$att{Owner} = ATT( $atp{Byte}, pack('H*','0000')); # PR_RCVD_REPRESENTING_xxx or PR_SENT_REPRESENTING_xxx
$att{SentFor} = ATT( $atp{Byte}, pack('H*','0001')); # PR_SENT_REPRESENTING_xxx
$att{Delegate} = ATT( $atp{Byte}, pack('H*','0002')); # PR_RCVD_REPRESENTING_xxx
$att{DateStart} = ATT( $atp{Date}, pack('H*','0006')); # PR_DATE_START
$att{DateEnd} = ATT( $atp{Date}, pack('H*','0007')); # PR_DATE_END
$att{AidOwner} = ATT( $atp{Long}, pack('H*','0008')); # PR_OWNER_APPT_ID
$att{RequestRes} = ATT( $atp{Short}, pack('H*','0009')); # PR_RESPONSE_REQUESTED

# Create reverse lookup table
while ( my ($att_name, $att_id) = each %att) {
 $att_name{$att_id}=$att_name;
}

$g_file_cnt=0;

# Set some package global defaults for new objects
# which can be overridden for any individual object.
%dflts = (
 debug=>0,
 debug_max_display=>1024,
 debug_max_line_size=>64,
 ignore_checksum=>0,
 display_after_err=>32,
 output_to_core=>4096,
 output_dir=>File::Spec->curdir,
 output_prefix=>"tnef",
 buffer_size=>1024,
);

sub mk_fname {
 my $parms = shift;
 File::Spec->catfile($parms->{output_dir},
                     $parms->{output_prefix}."-".$$."-".++$g_file_cnt.".doc");
}

sub rtn_err {
 my ($errmsg, $fh, $parms) = @_;
 $errstr = $errmsg;
 if ($parms->{debug}) {
  my $read_size = $parms->{display_after_err} || 32;
  my $data;
  $fh->read($data, $read_size);
  print "Error: $errstr\n";
  print "Data:\n";
  print $1,"\n"
   while $data =~ /([^\r\n]{0,$parms->{debug_max_line_size}})\r?\n?/g;
  print "HData:\n";
  my $hdata = unpack("H*",$data);
  print $1,"\n" while $hdata =~ /(.{0,$parms->{debug_max_line_size}})/g;
 }
 return undef;
}

sub read_err {
 my ($bytes, $fh, $errmsg) = @_;
 $errstr = (defined $bytes)? "Premature EOF" : "Read Error:".$errmsg;
 return undef;
}

sub equal {
 local $^W=0;
 $_[0] == $_[1];
}

# Autoload methods go after =cut, and are processed by the autosplit program.

1;
__END__

=head1 NAME

 Convert::TNEF - Perl module to read TNEF files

=head1 SYNOPSIS

 use Convert::TNEF;

 $tnef = Convert::TNEF->read($iohandle, \%parms)
  or die Convert::TNEF::errstr;

 $tnef = Convert::TNEF->read_in($filename, \%parms)
  or die Convert::TNEF::errstr;

 $tnef = Convert::TNEF->read_ent($mime_entity, \%parms)
  or die Convert::TNEF::errstr;

 $tnef->purge;

 $message = $tnef->message;

 @attachments = $tnef->attachments;

 $attribute_value      = $attachments[$i]->data($att_attribute_name);
 $attribute_value_size = $attachments[$i]->size($att_attribute_name);
 $attachment_name = $attachments[$i]->name;
 $long_attachment_name = $attachments[$i]->longname;

 $datahandle = $attachments[$i]->datahandle($att_attribute_name);

=head1 DESCRIPTION

 TNEF stands for Transport Network Encapsulation Format, and if you've
 ever been unfortunate enough to receive one of these files as an email
 attachment, you may want to use this module.

 read() takes as its first argument any file handle open
 for reading. The optional second argument is a hash reference
 which contains one or more of the following keys:

=head2

 output_dir - Path for storing TNEF attribute data kept in files
 (default: current directory).

 output_prefix - File prefix for TNEF attribute data kept in files
 (default: 'tnef').

 output_to_core - TNEF attribute data will be saved in core memory unless
 it is greater than this many bytes (default: 4096). May also be set to
 'NONE' to keep all data in files, or 'ALL' to keep all data in core.

 buffer_size - Buffer size for reading in the TNEF file (default: 1024).

 debug - If true, outputs all sorts of info about what the read() function
 is reading, including the raw ascii data along with the data converted
 to hex (default: false).

 display_after_err - If debug is true and an error is encountered,
 reads and displays this many bytes of data following the error
 (default: 32).

 debug_max_display - If debug is true then read and display at most
 this many bytes of data for each TNEF attribute (default: 1024).

 debug_max_line_size - If debug is true then at most this many bytes of
 data will be displayed on each line for each TNEF attribute
 (default: 64).

 ignore_checksum - If true, will ignore checksum errors while parsing
 data (default: false).

 read() returns an object containing the TNEF 'attributes' read from the
 file and the data for those attributes. If all you want are the
 attachments, then this is mostly garbage, but if you're interested then
 you can see all the garbage by turning on debugging. If the garbage
 proves useful to you, then let me know how I can maybe make it more
 useful.

 If an error is encountered, an undefined value is returned and the
 package variable $errstr is set to some helpful message.

 read_in() is a convienient front end for read() which takes a filename
 instead of a handle.

 read_ent() is another convient front end for read() which can take a
 MIME::Entity object (or any object with like methods, specifically
 open("r"), read($buff,$num_bytes), and close ).

 purge() deletes any on-disk data that may be in the attachments of
 the TNEF object.

 message() returns the message portion of the tnef object, if any.
 The thing it returns is like an attachment, but its not an attachment.
 For instance, it more than likely does not have a name or any
 attachment data.

 attachments() returns a list of the attachments that the given TNEF
 object contains.

 data() takes a TNEF attribute name, and returns a string value for that 
 attribute for that attachment. Its your own problem if the string is too
 big for memory. If no argument is given, then the 'AttachData' attribute
 is assumed, which is probably the data you're looking for.

 name() is the same as data(), except the attribute 'AttachTitle' is
 assumed, which returns the 8 character + 3 character extension name of
 the attachment.

 longname() returns the long filename and extension of an attachment. This
 is embedded in the 'Attachment' attribute data, so we attempt to extract
 the name out of that.

 size() takes an TNEF attribute name, and returns the size in bytes for
 the data for that attachment attribute.

 datahandle() is a method for attachments which takes a TNEF attribute
 name, and returns the data for that attribute as an object which is
 the same as a MIME::Body object.  See MIME::Body for all the applicable
 methods. If no argument is given, then 'AttachData' is assumed.


=head1 EXAMPLES

 # Here's a rather long example where mail is retrieved
 # from a POP3 server based on header information, then
 # it is MIME parsed, and then the TNEF contents
 # are extracted and converted.

 use strict;
 use Net::POP3;
 use MIME::Parser;
 use Convert::TNEF;

 # Redefine some functions
 # Hopefully these will be incorporated into the
 # regular distribution for these modules
 {
  local $^W=0;

  eval <<'EOT';
  use Carp;

  sub Net::POP3::get {
   @_ == 3 or croak 'usage: $pop3->get( MSGNUM, FH )';
   my ($me, $msg, $fh) = @_;

   return undef
     unless $me->_RETR($msg);

   $me->read_until_dot($fh);
  }

  sub Net::Cmd::read_until_dot {
   my $cmd = shift;
   my $fh  = shift;
   my $arr = [];

   while(1)
   {
    my $str = $cmd->getline() or return undef;

    $cmd->debug_print(0,$str)
      if ($cmd->debug & 4);

    last if($str =~ /^\.\r?\n/o);

    $str =~ s/^\.\././o;

    if ($fh) {
     print $fh $str;
    } else {
     push (@$arr, $str);
    }
   }

   $arr;
  }
 EOT
 }

 my $mail_dir = "mailout";
 my $mail_prefix = "mail";

 my $pop = new Net::POP3 ( "pop3server_name" );
 my $num_msgs = $pop->login("user_name","password");
 die "Can't login: $!" unless defined $num_msgs;

 # Get mail by sender and subject
 my $mail_out_idx = 0;
 MESSAGE: for ( my $i=1; $i<= $num_msgs;  $i++ ) {
  my $header = join "", @{$pop->top($i)};

  for ($header) {
   next MESSAGE unless
    /^from:.*someone\@somewhere.net/im &&
    /^subject:\s*important stuff/im
  }

  my $fname = $mail_prefix."-".$$.++$mail_out_idx.".doc";
  open (MAILOUT, ">$mail_dir/$fname")
   or die "Can't open $mail_dir/$fname: $!";
  $pop->get($i, \*MAILOUT) or die "Can't read mail";
  close MAILOUT or die "Error closing $mail_dir/$fname";
  # If you want to delete the mail on the server
  # $pop->delete($i);
 }

 close MAILOUT;
 $pop->quit();

 # Parse the mail message into separate mime entities
 my $parser=new MIME::Parser;
 $parser->output_dir("mimemail");

 opendir(DIR, $mail_dir) or die "Can't open directory $mail_dir: $!";
 my @files = map { $mail_dir."/".$_ } sort
  grep { -f "$mail_dir/$_" and /$mail_prefix-$$-/o } readdir DIR;
 closedir DIR;

 for my $file ( @files ) {
  my $entity=$parser->parse_in($file) or die "Couldn't parse mail";
  print_tnef_parts($entity);
  # If you want to delete the working files
  # $entity->purge;
 }

 sub print_tnef_parts {
  my $ent = shift;

  if ( $ent->parts ) {
   for my $sub_ent ( $ent->parts ) {
    print_tnef_parts($sub_ent);
   }
  } elsif ( $ent->mime_type =~ /ms-tnef/i ) {

   # Create a tnef object
   my $tnef = Convert::TNEF->read_ent($ent,{output_dir=>"tnefmail"})
   or die $Convert::TNEF::errstr;
   for ($tnef->attachments) {
    print "Title:",$_->name,"\n";
    print "Data:\n",$_->data,"\n";
   }

   # If you want to delete the working files
   # $tnef->purge;
  }
 }

=head1 SEE ALSO

perl(1), IO::Wrap(3), MIME::Parser(3), MIME::Entity(3), MIME::Body(3)

=head1 CAVEATS

 The parsing may depend on the endianness (see perlport) and width of
 integers on the system where the TNEF file was created. If this proves
 to be the case (check the debug output), I'll see what I can do
 about it.

=head1 AUTHOR

 Douglas Wilson, dwilson@gtemail.net

=cut

sub read_ent {
 croak "Usage: Convert::TNEF->read_ent(entity, parameters) "
  unless @_ == 2 or @_ == 3;
 my $self = shift;
 my ($ent, $parms) = @_;
 my $io=$ent->open("r") or do {
  $errstr = "Can't open entity: $!";
  return undef;
 };
 my $tnef=$self->read($io, $parms);
 $io->close or do {
  $errstr = "Error closing handle: $!";
  return undef;
 };
 return $tnef;
}

sub read_in {
 croak "Usage: Convert::TNEF->read_in(filename, parameters) "
  unless @_ == 2 or @_ == 3;
 my $self = shift;
 my ($fname, $parms) = @_;
 open(INFILE, "<$fname") or do {
   $errstr = "Can't open $fname: $!";
   return undef;
 };
 binmode INFILE;
 my $tnef = $self->read(\*INFILE, $parms);
 close INFILE or do {
  $errstr = "Error closing $fname: $!";
  return undef;
 };
 return $tnef;
}

sub read {
 croak "Usage: Convert::TNEF->read(fh, parameters) "
  unless @_ == 2 or @_ == 3;
 my $self = shift;
 my $class = ref($self) || $self;
 $self = {};
 bless $self, $class;
 my ($fd, $parms) = @_;
 $fd = wraphandle($fd);

 my %parms = %dflts;
 @parms{keys %$parms} = values %$parms if defined $parms;
 $parms = \%parms;
 my $debug = $parms{debug};
 my $ignore_checksum = $parms{ignore_checksum};
 
 # Start of TNEF stream
 my $data;
 my $num_bytes = $fd->read($data, 4);
 return read_err($num_bytes,$fd,$!) unless equal($num_bytes,4);
 print "TNEF start: ",unpack("H*", $data),"\n" if $debug;
 return rtn_err("Not TNEF-encapsulated", $fd, $parms)
  unless $data eq $TNEF_SIGNATURE;

 # Key
 $num_bytes = $fd->read($data, 2);
 return read_err($num_bytes,$fd,$!) unless equal($num_bytes,2);
 print "TNEF key: ",unpack("H*", $data),"\n" if $debug;

 # Start of First Object
 $num_bytes = $fd->read($data, 1);
 return read_err($num_bytes,$fd,$!) unless equal($num_bytes,1);

 my $msg_att="";
 my $got_version = 0;
 my $att_cnt=0;

 my $is_msg = ($data eq $LVL_MESSAGE);
 my $is_att = ($data eq $LVL_ATTACHMENT);
 print "TNEF object start: ",unpack("H*", $data),"\n" if $debug;
 return rtn_err("Neither a message or an attachment", $fd, $parms)
  unless $is_msg or $is_att;

 my %msg=();
 my %msg_size=();
 my $is_eof = 0;
 while ($is_msg) {
  print "Message attribute\n" if $is_msg and $debug;
  $num_bytes = $fd->read($data, 4);
  return read_err($num_bytes,$fd,$!) unless equal($num_bytes,4);
  my $att_id = $data;

  print "TNEF message attribute: ",unpack("H*", $data),"\n" if $debug;
  return rtn_err("Bad Attribute found in Message", $fd, $parms)
   unless $att_name{$att_id};
  if ($att_id eq $att{TnefVersion}) {
   return rtn_err("Version must be first attribute in message", $fd, $parms)
    if $att_cnt;
   return rtn_err("Two attTnefVersion records found", $fd, $parms)
    if $got_version;
   $got_version = 1;
  }
  $att_cnt++;
  print "Got attribute:$att_name{$att_id}\n" if $debug;

  $num_bytes = $fd->read($data, 4);
  return read_err($num_bytes,$fd,$!) unless equal($num_bytes,4);

  print "HLength:", unpack("H8",$data), "\n" if $debug;
  my $length = unpack("V", $data);
  print "Length: $length\n" if $debug;

  # Get the attribute data (returns an object since data may
  # actually end up in a file)
  my $calc_chksum;
  $data=$self->build_data($fd, $length, \$calc_chksum, $parms) or return undef;
  $self->debug_print($length, $att_id, $data, $parms) if $debug;
  $msg{$att_name{$att_id}} = $data;

  $num_bytes = $fd->read($data, 2);
  return read_err($num_bytes,$fd,$!) unless equal($num_bytes,2);
  my $file_chksum = $data;
  if ($debug) {
   print "Calc Chksum:", unpack("H*",$calc_chksum),"\n";
   print "File Chksum:", unpack("H*",$file_chksum),"\n";
  }
  return rtn_err("Bad Checksum", $fd, $parms)
   unless $calc_chksum eq $file_chksum or $ignore_checksum;

  my $num_bytes=$fd->read($data, 1);
  # EOF (0 bytes) is ok
  return read_err($num_bytes,$fd,$!) unless defined $num_bytes;
  print "Next token:", unpack("H2", $data), "\n" if $debug;
  $is_msg = ($data eq $LVL_MESSAGE);
  $is_att = ($data eq $LVL_ATTACHMENT);
  $is_eof = ($num_bytes == 0);
 }

 return rtn_err("Not an attachment", $fd, $parms) unless $is_eof or $is_att;

 my @atts=();
 my @atts_size=();
 my %attachment=();
 my %attachment_size=();
 while ($is_att) {
  print "Attachment attribute\n" if $debug;

  $num_bytes = $fd->read($data, 4);
  return read_err($num_bytes,$fd,$!) unless equal($num_bytes,4);
  my $att_id = $data;

  print "TNEF attachment attribute: ",unpack("H*", $data),"\n" if $debug;
  return rtn_err("Bad Attribute found in attachment", $fd, $parms)
   unless $att_name{$att_id};
  return rtn_err("Version record found in attachment", $fd, $parms)
   if $att_id eq $att{TnefVersion};
  print "Got attribute:$att_name{$att_id}\n" if $debug;

  $num_bytes = $fd->read($data, 4);
  return read_err($num_bytes,$fd,$!) unless equal($num_bytes,4);

  print "HLength:", unpack("H8",$data), "\n" if $debug;
  my $length = unpack("V", $data);
  print "Length: $length\n" if $debug;

  # Get the attribute data (returns an object since data may
  # actually end up in a file)
  my $calc_chksum;
  $data=$self->build_data($fd, $length, \$calc_chksum, $parms) or return undef;
  $self->debug_print($length, $att_id, $data, $parms) if $debug;

  # See if we're starting a new attachment
  if ($att_id eq $att{AttachRenddata}) {
   # Save the previous attachment, if any
   if (%attachment) {
    $attachment{TN_Size} = {%attachment_size};
    my $attref = {%attachment};
    bless $attref, $class;
    push(@atts, $attref);
   }
   %attachment = (AttachRendData=>$data);
   %attachment_size = (AttachRendData=>$length);
  } else {
   $attachment{$att_name{$att_id}} = $data;
   $attachment_size{$att_name{$att_id}} = $length;
  }

  # Validate checksum
  $num_bytes = $fd->read($data, 2);
  return read_err($num_bytes,$fd,$!) unless equal($num_bytes,2);
  my $file_chksum = $data;
  print "Calc Chksum:", unpack("H*",$calc_chksum),"\n" if $debug;
  print "File Chksum:", unpack("H*",$file_chksum),"\n" if $debug;
  die "Bad Checksum\n" unless $calc_chksum eq $file_chksum or $ignore_checksum;

  my $num_bytes=$fd->read($data, 1);
  # EOF (0 bytes) is ok
  return read_err($num_bytes,$fd,$!) unless defined $num_bytes;
  print "Next token:", unpack("H2", $data), "\n" if $debug;
  $is_msg = ($data eq $LVL_MESSAGE);
  $is_att = ($data eq $LVL_ATTACHMENT);
  $is_eof = ($num_bytes < 1);
 }
 if (%attachment) {
  my $attref = {%attachment};
  my $attsiz = {%attachment_siz};
  bless $attref, $class;
  push(@atts, $attref);
  push(@atts_size, $attsiz);
 }

 return rtn_err("Message found in attachment section",$fd,$parms)
  if $is_msg;
 return rtn_err("Not an attachment", $fd, $parms) unless $is_eof or $is_att;
 print "EOF\n" if $debug and $is_eof;

 $msg{TN_Size} = \%msg_size;
 my $msg = \%msg;
 bless $msg, $class;
 $self->{TN_Message} = $msg;
 $self->{TN_Attachments} = \@atts;
 return $self;
}

sub debug_print {
 my ($self, $length, $att_id, $data, $parms) = @_;
 if ($length < $parms->{debug_max_display}) {
  $data = $data->data;
  if ($att_id eq $att{TnefVersion}) {
   $data = unpack("L", $data);
   print "Version: $data\n";
  } elsif (substr($att_id, 2) eq $atp{Date} and $length == 14) {
   my ($yr, $mo, $day, $hr, $min, $sec, $dow) = unpack("vvvvvvv", $data);
   my $date = join ":", $yr, $mo, $day, $hr, $min, $sec, $dow;
   print "Date: $date\n";
   print "HDate:", unpack("H*",$data), "\n";
  } elsif ($att_id eq $att{AttachRenddata} and $length == 14) {
   my ($atyp,$ulPosition,$dxWidth,$dyHeight,$dwFlags) = unpack("vVvvV", $data);
   $data = join ":", $atyp, $ulPosition, $dxWidth, $dyHeight, $dwFlags;
   print "AttachRendData: $data\n";
  } else {
   print "Data:\n";
   print $1,"\n"
    while $data =~ /([^\r\n]{0,$parms->{debug_max_line_size}})\r?\n?/g;
   print "HData:\n";
   my $hdata = unpack("H*",$data);
   print $1,"\n" while $hdata =~ /(.{0,$parms->{debug_max_line_size}})/g;
  }
 } else {
  my $io=$data->open("r") or croak "Error opening attachment data handle: $!";
  my $buffer;
  $io->read($buffer, $parms->{debug_max_display});
  $io->close or croak "Error closing attachment data handle: $!";
  print "Data:\n";
  print $1,"\n"
   while $buffer =~ /([^\r\n]{0,$parms->{debug_max_line_size}})\r?\n?/sg;
  print "HData:\n";
  my $hdata = unpack("H*",$buffer);
  print $1,"\n" while $hdata =~ /(.{0,$parms->{debug_max_line_size}})/g;
 }
}

sub build_data {
 my ($self, $fd, $length, $chksumref, $parms) = @_;
 my $cutoff = $parms->{output_to_core};
 my $incore = do {
  if    ($cutoff eq 'NONE') { 0 } #Everything to files
  elsif ($cutoff eq 'ALL') { 1 }  #Everything in memory
  elsif ($cutoff < $length) { 0 } #Large items in files
  else { 1 }                      #Everything else in memory
 };

 my $body = ($incore)? new Convert::TNEF::Scalar
                     : new Convert::TNEF::File mk_fname($parms);
 $body->binmode(1);
 my $io = $body->open("w");
 my $bufsiz = $parms->{buffer_size};
 $bufsiz = $length if $length < $bufsiz;
 my $buffer;
 my $chksum = 0;
 while ($length > 0) {
  my $num_bytes=$fd->read($buffer, $bufsiz);
  return read_err($num_bytes,$fd,$!) unless equal($num_bytes,$bufsiz);
  $io->print($buffer);
  $chksum += unpack("%16C*", $buffer);
  $chksum %= 65536;
  $length -= $bufsiz;
  $bufsiz = $length if $length < $bufsiz;
 }
 $$chksumref = pack("v", $chksum);
 $io->close;
 return $body;
}

sub purge {
 my $self = shift;
 for ($self->attachments) {
  $_->purge;
 }
}

sub purge {
 my $self = shift;
 my $msg = $self->{TN_Message};
 my @atts = $self->attachments;
 for (keys %$msg) {
  $msg->{$_}->purge if exists $att{$_};
 }
 for my $attch (@atts) {
  for (keys %$attch) {
   $attch->{$_}->purge if exists $att{$_};
  }
 }
}

sub message {
 $self->{TN_Message};
}

sub attachments {
 my $self = shift;
 my $attachments = $self->{TN_Attachments} || [];
 return @{$attachments};
}

sub data {
 my $self = shift;
 my $attr = shift || 'AttachData';
 return $self->{$attr}->data;
}

sub name {
 my $self = shift;
 my $attr = shift || 'AttachTitle';
 return $self->{$attr}->data;
}

sub longname {
 my $self = shift;
 my $name = $self->name;
 my ($prfx, $ext) = $name =~ /(.{6})~\d(\..{0,3})/;
 return $name unless defined $prfx;

 my $data = $self->data("Attachment");
 my ($longname) = $data =~ /\x00(\Q$prfx\E[^\x00]+\Q$ext\E)\x00/i;
 return $longname || $name;

}

sub datahandle {
 my $self = shift;
 my $attr = shift || 'AttachData';
 return $self->{$attr};
}

sub size {
 my $self = shift;
 my $attr = shift || 'AttachData';
 return $self->{TN_Size}->{$attr};
}
