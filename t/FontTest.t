=begin comment

Copyright (c) 2019 Aspose Pty Ltd

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

=end comment

=cut

use File::Slurp;
use JSON;

use Test::More;
use Test::Exception;

use AsposeSlidesCloud::TestUtils;
use AsposeSlidesCloud::SlidesApi;
use AsposeSlidesCloud::SlidesAsyncApi;
use AsposeSlidesCloud::Object::ImageExportOptions;
use AsposeSlidesCloud::Object::FontSubstRule;

use strict;
use warnings;

my $utils = AsposeSlidesCloud::TestUtils->new();

subtest 'get available fonts' => sub {
    eval {
        my $response = $utils->{testSlidesApi}->get_available_fonts();
        ok(scalar @{$response->{list}} > 1);
        ok(!$response->{list}[0]->{is_custom});
    };
    if ($@) {
        fail("get_available_fonts raised an exception: $@");
    }
    pass();
};

subtest 'get fonts' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->get_fonts(%params);
        is(scalar @{$response->{list}}, 4);
    };
    if ($@) {
        fail("get_fonts raised an exception: $@");
    }
    pass();
};

subtest 'get fonts online' => sub {
    eval {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my %params = ('document' => $source, 'password' => "password");
        my $response = $utils->{testSlidesApi}->get_fonts_online(%params);
        is(scalar @{$response->{list}}, 4);
    };
    if ($@) {
        fail("get_fonts_online raised an exception: $@");
    }
    pass();
};

subtest 'set embedded font' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'font_name' => "Calibri", 'only_used' => 0, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->set_embedded_font(%params);
        is(scalar @{$response->{list}}, 4);
        ok(!$response->{list}[0]->{is_embedded});
        ok($response->{list}[1]->{is_embedded});
        is($response->{list}[2]->{font_name}, "Calibri");
        ok($response->{list}[2]->{is_embedded});
    };
    if ($@) {
        fail("set_embedded_font raised an exception: $@");
    }
    pass();
};

subtest 'set embedded font online' => sub {
    eval {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my %params = ('document' => $source, 'font_name' => "Calibri", 'only_used' => 0, 'password' => "password");
        my $response = $utils->{testSlidesApi}->set_embedded_font_online(%params);
        ok(length($response) > length($source));
    };
    if ($@) {
        fail("set_embedded_font_online raised an exception: $@");
    }
    pass();
};

subtest 'set embedded fonts' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my @font_names = ("Calibri");
        my %params = ('name' => "test.pptx", 'font_names' => \@font_names, 'only_used' => 0, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->set_embedded_fonts(%params);
        is(scalar @{$response->{list}}, 4);
        ok(!$response->{list}[0]->{is_embedded});
        ok($response->{list}[1]->{is_embedded});
        is($response->{list}[2]->{font_name}, "Calibri");
        ok($response->{list}[2]->{is_embedded});
    };
    if ($@) {
        fail("set_embedded_fonts raised an exception: $@");
    }
    pass();
};

subtest 'set embedded fonts from request' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $font_file = read_file("TestData/calibri.ttf", { binmode => ':raw' });
        my @fonts = ($font_file);
        my %params = ('name' => "test.pptx", 'fonts' => \@fonts, 'only_used' => 0, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->set_embedded_fonts(%params);
        is(scalar @{$response->{list}}, 4);
        ok(!$response->{list}[0]->{is_embedded});
        ok($response->{list}[1]->{is_embedded});
        is($response->{list}[2]->{font_name}, "Calibri");
        ok($response->{list}[2]->{is_embedded});
    };
    if ($@) {
        fail("set_embedded_fonts raised an exception: $@");
    }
    pass();
};

subtest 'set embedded fonts online' => sub {
    eval {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my @font_names = ("Calibri");
        my %params = ('document' => $source, 'font_names' => \@font_names, 'only_used' => 0, 'password' => "password");
        my $response = $utils->{testSlidesApi}->set_embedded_fonts_online(%params);
        ok(length($response) > length($source));
    };
    if ($@) {
        fail("set_embedded_fonts_online raised an exception: $@");
    }
    pass();
};

subtest 'set embedded fonts from request online' => sub {
    eval {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my $font_file = read_file("TestData/calibri.ttf", { binmode => ':raw' });
        my @fonts = ($font_file);
        my %params = ('document' => $source, 'fonts' => \@fonts, 'only_used' => 0, 'password' => "password");
        my $response = $utils->{testSlidesApi}->set_embedded_fonts_online(%params);
        ok(length($response) > length($source));
    };
    if ($@) {
        fail("set_embedded_fonts_online raised an exception: $@");
    }
    pass();
};

subtest 'compress embedded fonts' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'font_name' => "Calibri", 'only_used' => 0, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->set_embedded_font(%params);
        ok($response->{list}[2]->{is_embedded});

        %params = ('name' => "test.pptx", 'password' => "password", 'folder' => "TempSlidesSDK");
        $utils->{testSlidesApi}->compress_embedded_fonts(%params);
    };
    if ($@) {
        fail("compress_embedded_fonts raised an exception: $@");
    }
    pass();
};

subtest 'compress embedded fonts online' => sub {
    eval {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my %params = ('document' => $source, 'font_name' => "Calibri", 'only_used' => 0, 'password' => "password");
        my $response_embedded = $utils->{testSlidesApi}->set_embedded_font_online(%params);
        ok(length($response_embedded) > length($source));

        %params = ('document' => $response_embedded, 'password' => "password");
        my $response_compressed = $utils->{testSlidesApi}->compress_embedded_fonts_online(%params);
        ok(length($response_compressed) < length($response_embedded));
    };
    if ($@) {
        fail("compress_embedded_fonts_online raised an exception: $@");
    }
    pass();
};

subtest 'delete embedded font' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'font_name' => "Calibri", 'only_used' => 0, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->set_embedded_font(%params);
        ok($response->{list}[2]->{is_embedded});

        %params = ('name' => "test.pptx", 'font_name' => "Calibri", 'password' => "password", 'folder' => "TempSlidesSDK");
        $response = $utils->{testSlidesApi}->delete_embedded_font(%params);
        ok(!$response->{list}[2]->{is_embedded});
    };
    if ($@) {
        fail("delete_embedded_font raised an exception: $@");
    }
    pass();
};

subtest 'delete embedded font online' => sub {
    eval {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my %params = ('document' => $source, 'font_name' => "Calibri", 'only_used' => 0, 'password' => "password");
        my $response_embedded = $utils->{testSlidesApi}->set_embedded_font_online(%params);
        ok(length($response_embedded) > length($source));

        %params = ('document' => $response_embedded, 'font_name' => "Calibri", 'password' => "password");
        my $response_deleted = $utils->{testSlidesApi}->delete_embedded_font_online(%params);
        ok(length($response_deleted) < length($response_embedded));
    };
    if ($@) {
        fail("delete_embedded_font_online raised an exception: $@");
    }
    pass();
};

subtest 'replace font' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'source_font' => "Calibri", 'target_font' => "Times New Roman", 'embed' => 1, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->replace_font(%params);
        is($response->{list}[2]->{font_name}, "Times New Roman");
        ok($response->{list}[2]->{is_embedded});
    };
    if ($@) {
        fail("replace_font raised an exception: $@");
    }
    pass();
};

subtest 'replace font online' => sub {
    eval {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my %params = ('document' => $source, 'source_font' => "Calibri", 'target_font' => "Times New Roman", 'embed' => 1, 'password' => "password");
        my $response = $utils->{testSlidesApi}->replace_font_online(%params);
        ok(length($response) > length($source));
    };
    if ($@) {
        fail("replace_font_online raised an exception: $@");
    }
    pass();
};

subtest 'font substitution' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $rule1 = AsposeSlidesCloud::Object::FontSubstRule->new();
        $rule1->{source_font} = "Arial";
        $rule1->{target_font} = "Times New Roman";
        $rule1->{not_found_only} = 0;
        my $rule2 = AsposeSlidesCloud::Object::FontSubstRule->new();
        $rule2->{source_font} = "Calibri";
        $rule2->{target_font} = "Times New Roman";
        my @rules = ($rule1, $rule2);

        my $export_options = AsposeSlidesCloud::Object::ImageExportOptions->new();
        $export_options->{font_subst_rules} = \@rules;

        my %params = ('name' => "test.pptx", 'format' => 'Png', 'options' => $export_options, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->download_presentation(%params);
        ok(length($response) > 0);
    };
    if ($@) {
        fail("download_presentation raised an exception: $@");
    }
    pass();
};

done_testing;
