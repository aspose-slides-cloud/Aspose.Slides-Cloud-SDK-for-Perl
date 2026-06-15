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
use Archive::Zip;
use JSON;

use Test::More;
use Test::Exception;

use AsposeSlidesCloud::TestUtils;
use AsposeSlidesCloud::SlidesApi;
use AsposeSlidesCloud::SlidesAsyncApi;

use strict;
use warnings;

my $utils = AsposeSlidesCloud::TestUtils->new();

subtest 'images get' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'password' => "password", 'folder' => "TempSlidesSDK");
        my $presentation_result = $utils->{testSlidesApi}->get_presentation_images(%params);

        $params{slide_index} = 1;
        my $slide_result = $utils->{testSlidesApi}->get_slide_images(%params);
        ok(scalar @{$slide_result->{list}} < @{$presentation_result->{list}});

        $params{slide_index} = 2;
        $params{shape_index} = 2;
        my $shape_result = $utils->{testSlidesApi}->get_slide_images(%params);
        ok(scalar @{$shape_result->{list}} == 1);

        undef $params{shape_index};
        $params{alt_text} = "title";
        my $alt_result = $utils->{testSlidesApi}->get_slide_images(%params);
        ok(scalar @{$alt_result->{list}} == 0);
    };
    if ($@) {
        fail("get_presentation_images raised an exception: $@");
    }
    pass();
};

subtest 'images download all storage' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'password' => "password", 'folder' => "TempSlidesSDK");
        my $default_result = $utils->{testSlidesApi}->download_images_default_format(%params);

        $params{format} = "png";
        my $png_result = $utils->{testSlidesApi}->download_images(%params);

        my $defaultFile = 'defaultZip.zip';
	open my $fh, '>>', $defaultFile;
	binmode $fh;
	print $fh $default_result;
	close $fh;

        my $defaultZip = Archive::Zip->new();
        $defaultZip->read($defaultFile);
        my @defaultFiles = $defaultZip->memberNames();

        my $pngFile = 'pngZip.zip';
	open $fh, '>>', $pngFile;
	binmode $fh;
	print $fh $png_result;
	close $fh;

        my $pngZip = Archive::Zip->new();
        $pngZip->read($pngFile);
        my @pngFiles = $pngZip->memberNames();

        is(scalar @defaultFiles, scalar @pngFiles);
    };
    if ($@) {
        fail("download_images raised an exception: $@");
    }
    pass();
};

subtest 'images download all request' => sub {
    eval {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my %params = ('document' => $source, 'password' => "password");
        my $default_result = $utils->{testSlidesApi}->download_images_default_format_online(%params);

        $params{format} = "png";
        my $png_result = $utils->{testSlidesApi}->download_images_online(%params);

        my $defaultFile = 'defaultZip.zip';
	open my $fh, '>>', $defaultFile;
	binmode $fh;
	print $fh $default_result;
	close $fh;

        my $defaultZip = Archive::Zip->new();
        $defaultZip->read($defaultFile);
        my @defaultFiles = $defaultZip->memberNames();

        my $pngFile = 'pngZip.zip';
	open $fh, '>>', $pngFile;
	binmode $fh;
	print $fh $png_result;
	close $fh;

        my $pngZip = Archive::Zip->new();
        $pngZip->read($pngFile);
        my @pngFiles = $pngZip->memberNames();

        is(scalar @defaultFiles, scalar @pngFiles);
    };
    if ($@) {
        fail("download_images raised an exception: $@");
    }
    pass();
};

subtest 'images download storage' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my %params = ('name' => "test.pptx", 'index' => 1, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $default_result = $utils->{testSlidesApi}->download_image_default_format(%params);

        $params{format} = "png";
        my $png_result = $utils->{testSlidesApi}->download_image(%params);

        ok(not length($default_result) == length($png_result));
    };
    if ($@) {
        fail("download_image raised an exception: $@");
    }
    pass();
};

subtest 'images download request' => sub {
    eval {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my %params = ('document' => $source, 'index' => 1, 'password' => "password");
        my $default_result = $utils->{testSlidesApi}->download_image_default_format_online(%params);

        $params{format} = "png";
        my $png_result = $utils->{testSlidesApi}->download_image_online(%params);

        ok(not length($default_result) == length($png_result));
    };
    if ($@) {
        fail("download_image raised an exception: $@");
    }
    pass();
};

subtest 'compress image' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'slide_index' => 2, 'shape_index' => 2,, 'resolution' => 150, 'password' => "password", 'folder' => "TempSlidesSDK");
        $utils->{testSlidesApi}->compress_image(%params);
    };
    if ($@) {
        fail("compress_image raised an exception: $@");
    }
    pass();
};

subtest 'delete picture cropped areas' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'slide_index' => 2, 'shape_index' => 2, 'delete_picture_cropped_areas' => 1, 'password' => "password", 'folder' => "TempSlidesSDK");
        $utils->{testSlidesApi}->compress_image(%params);
    };
    if ($@) {
        fail("compress_image raised an exception: $@");
    }
    pass();
};

subtest 'delete picture cropped areas wrong shape type' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'slide_index' => 2, 'shape_index' => 3, 'delete_picture_cropped_areas' => 1, 'password' => "password", 'folder' => "TempSlidesSDK");
        $utils->{testSlidesApi}->compress_image(%params);
    };
    if ($@) {
        if ($@ =~ m/API Exception\((\d+)\): (.*) /s) {
            is($1, 400);
        } else {
            fail("Invalid exception type: $@");
        }
    } else {
        fail("Should throw an exception if shape is not PictureFrame");
    }
    pass();
};

subtest 'image download quality' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'index' => 1, 'format' => 'Jpeg', 'password' => "password", 'folder' => "TempSlidesSDK", 'quality' => 100);
        my $image_good = $utils->{testSlidesApi}->download_image(%params);

        $params{quality} = 50;
        my $image_bad = $utils->{testSlidesApi}->download_image(%params);

        ok(length($image_good) > length($image_bad));
    };
    if ($@) {
        fail("download_image raised an exception: $@");
    }
    pass();
};

subtest 'image download quality useless' => sub {
    eval {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my %params = ('document' => $source, 'index' => 1, 'format' => 'Png', 'password' => "password", 'quality' => 100);
        my $image_good = $utils->{testSlidesApi}->download_image_online(%params);

        $params{quality} = 50;
        my $image_bad = $utils->{testSlidesApi}->download_image_online(%params);

        is(length($image_good), length($image_bad));
    };
    if ($@) {
        fail("download_image_online raised an exception: $@");
    }
    pass();
};

subtest 'replace image' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $image = read_file("TestData/watermark.png", { binmode => ':raw' });
        my %params = ('name' => "test.pptx", 'image_index' => 1, 'image' => $image, 'password' => "password", 'folder' => "TempSlidesSDK");
        $utils->{testSlidesApi}->replace_image(%params);
    };
    if ($@) {
        fail("replace_image raised an exception: $@");
    }
    pass();
};

subtest 'replace image request' => sub {
    eval {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my $image = read_file("TestData/watermark.png", { binmode => ':raw' });
        my %params = ('document' => $source, 'image_index' => 1, 'image' => $image, 'password' => "password");
        my $response = $utils->{testSlidesApi}->replace_image_online(%params);
        ok(length($response) > 0);
    };
    if ($@) {
        fail("replace_image_online raised an exception: $@");
    }
    pass();
};

done_testing;
