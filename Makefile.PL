use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    'NAME'         => 'Convert::TNEF',
    'VERSION_FROM' => 'TNEF.pm', # finds $VERSION
    'PREREQ_PM'    => { 'IO::Wrap'   => 1.103,
                        'MIME::Body' => 4.109,
                      },
    'dist'         => {COMPRESS => 'gzip',
                       SUFFIX   => 'gz',
                      },
);
