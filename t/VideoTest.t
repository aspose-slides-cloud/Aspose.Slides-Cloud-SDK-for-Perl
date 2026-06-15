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
use AsposeSlidesCloud::Object::VideoFrame;

use strict;
use warnings;

my $utils = AsposeSlidesCloud::TestUtils->new();

subtest 'video caption tracks' => sub {
    eval {
        my $slide_index = 3;
        my $shape_index = 3;
        my $track1_label = "track1";
        my $track2_label = "track2";
        my $track1_data = "WEBVTT\n\n00:00:00.000 --> 00:00:10.000\nCaption 1 text.";
        my $track2_data = "WEBVTT\n\n00:00:00.000 --> 00:00:10.000\nCaption 2 text.";

        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::VideoFrame->new();
        $dto->{base64_data} = "bXAzc2FtcGxl";
        my %create_params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
        $utils->{testSlidesApi}->create_shape(%create_params);

        my %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape_index, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $captions = $utils->{testSlidesApi}->get_caption_tracks(%params);
        is(scalar @{$captions->{items}}, 0);

        %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape_index, 'label' => $track1_label, 'data' => $track1_data, 'password' => "password", 'folder' => "TempSlidesSDK");
        $utils->{testSlidesApi}->create_caption_track(%params);
        %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape_index, 'label' => $track2_label, 'data' => $track2_data, 'password' => "password", 'folder' => "TempSlidesSDK");
        $utils->{testSlidesApi}->create_caption_track(%params);

        %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape_index, 'include_data' => 1, 'password' => "password", 'folder' => "TempSlidesSDK");
        $captions = $utils->{testSlidesApi}->get_caption_tracks(%params);
        is(scalar @{$captions->{items}}, 2);
        is($captions->{items}[0]->{label}, $track1_label);
        is($captions->{items}[1]->{label}, $track2_label);

        %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape_index, 'captions_index' => 1, 'password' => "password", 'folder' => "TempSlidesSDK");
        $utils->{testSlidesApi}->delete_caption_track(%params);

        %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape_index, 'include_data' => 0, 'password' => "password", 'folder' => "TempSlidesSDK");
        $captions = $utils->{testSlidesApi}->get_caption_tracks(%params);
        is(scalar @{$captions->{items}}, 1);
        is($captions->{items}[0]->{label}, $track2_label);

        %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape_index, 'password' => "password", 'folder' => "TempSlidesSDK");
        $utils->{testSlidesApi}->delete_caption_tracks(%params);

        %params = ('name' => "test.pptx", 'slide_index' => $slide_index, 'shape_index' => $shape_index, 'include_data' => 1, 'password' => "password", 'folder' => "TempSlidesSDK");
        $captions = $utils->{testSlidesApi}->get_caption_tracks(%params);
        is(scalar @{$captions->{items}}, 0);
    };
    if ($@) {
        fail("video caption tracks raised an exception: $@");
    }
    pass();
};

done_testing;
