# Before `make install' is performed this script should be runnable with
# `make test'. After `make install' it should work as `perl test.pl'

######################### We start with some black magic to print on failure.

# Change 1..1 below to 1..last_test_to_print .
# (It may become useful if the test is moved to ./t subdirectory.)

BEGIN { $| = 1; print "1..5\n"; }
END {print "not ok 1\n" unless $loaded;}
use Convert::TNEF;
$loaded = 1;
print "ok 1\n";

######################### End of black magic.

# Insert your test code below (better if it prints "ok 13"
# (correspondingly "not ok 13") depending on the success of chunk 13
# of the test code):

use strict;
use Convert::TNEF;

my $n = 2;
my $tnef = Convert::TNEF->read_in("t/tnef.doc",{ignore_checksum=>1});
if (ref($tnef) eq 'Convert::TNEF') {
 print "ok $n\n";
} else {
 print "not ok $n\n";
 exit;
}

$n++;
my $i=1;
my $t_cnt=0;
my $t_err=0;
my $d_cnt=0;
my $d_err=0;
for my $title ($tnef->attachments) {
 if ($title ne "tmp.out" or $t_cnt > 1) {
  $t_err++;
  last;
 } else {
  $t_cnt++;
 }
 for my $data ($tnef->data($title)) {
  if ($data !~ /^This is an attachment/ or $d_cnt > 1) {
   $d_err++;
   last;
  } else {
   $d_cnt++;
  }
 }
}

if ($t_err or not $t_cnt) {
 print "ok $n\n";
} else {
 print "not ok $n\n";
}

$n++;
if ($d_err or not $d_cnt) {
 print "ok $n\n";
} else {
 print "not ok $n\n";
}

$n++;
$tnef = Convert::TNEF->read_in("t/tnef.doc");
if ($Convert::TNEF::errstr eq 'Bad Checksum') {
 print "ok $n\n";
} else {
 print "not ok $n\n";
}
