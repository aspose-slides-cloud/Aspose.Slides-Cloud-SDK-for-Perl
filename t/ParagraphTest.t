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

BEGIN {push @INC, "../lib"}

use File::Slurp;
use Archive::Zip;
use JSON;

use Test::More;
use Test::Exception;

use AsposeSlidesCloud::TestUtils;
use AsposeSlidesCloud::SlidesApi;
use AsposeSlidesCloud::SlidesAsyncApi;
use AsposeSlidesCloud::Object::Shape;
use AsposeSlidesCloud::Object::Paragraph;
use AsposeSlidesCloud::Object::Portion;
use AsposeSlidesCloud::Object::PortionFormat;

use strict;
use warnings;

my $utils = AsposeSlidesCloud::TestUtils->new();

subtest 'get paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = ('name' => "test.pptx", 'slide_index' => 6, 'shape_index' => 2, 'paragraph_index' => 1, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->get_paragraph(%params);
        is(scalar @{$response->{portion_list}}, 2);
	};
    if ($@) {
        fail("get_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'get paragraphs' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = ('name' => "test.pptx", 'slide_index' => 6, 'shape_index' => 2, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->get_paragraphs(%params);
        is(scalar @{$response->{paragraph_links}}, 2);
	};
    if ($@) {
        fail("get_paragraphs raised an exception: $@");
    }
    pass();
};

subtest 'get sub-shape paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 3, 
            'paragraph_index' => 1, 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'sub_shape' => '1');

        my $response = $utils->{testSlidesApi}->get_paragraph(%params);
        is(scalar @{$response->{portion_list}}, 2);
	};
    if ($@) {
        fail("get_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'get sub-shape paragraphs' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 3, 
            'password' => "password", 
            'folder' => "TempSlidesSDK",
            "sub_shape" => '1');
        my $response = $utils->{testSlidesApi}->get_paragraphs(%params);
        is(scalar @{$response->{paragraph_links}}, 2);
	};
    if ($@) {
        fail("get_paragraphs raised an exception: $@");
    }
    pass();
};

subtest 'create paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::Paragraph->new();
        $dto->{margin_left} = 2;
        $dto->{margin_right} = 2;
        $dto->{alignment} = 'Center';
        
		my %params = ('name' => "test.pptx", 'slide_index' => 6, 'shape_index' => 2, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->create_paragraph(%params);
        
        is($response->{margin_left}, 2);
        is($response->{margin_right}, 2);
        is($response->{alignment}, 'Center');
	};
    if ($@) {
        fail("create_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'create paragraph width portions' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::Paragraph->new();
        my $portion1 = AsposeSlidesCloud::Object::Portion->new();
        $portion1->{text} = 'Portion1';
        my $portion2 = AsposeSlidesCloud::Object::Portion->new();
        $portion2->{text} = 'Portion2';
        my @portions = ($portion1, $portion2);
        $dto->{portion_list} = \@portions;
        
		my %params = ('name' => "test.pptx", 'slide_index' => 6, 'shape_index' => 2, 'dto' => $dto, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->create_paragraph(%params);
        
        is(scalar @{$response->{portion_list}}, 2);
	};
    if ($@) {
        fail("create_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'create sub-shape paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::Paragraph->new();
        $dto->{margin_left} = 2;
        $dto->{margin_right} = 2;
        $dto->{alignment} = 'Center';
        
		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 3, 
            'dto' => $dto, 
            'password' => "password", 
            'folder' => "TempSlidesSDK",
            'sub_shape' => '1');
        my $response = $utils->{testSlidesApi}->create_paragraph(%params);
        
        is($response->{margin_left}, 2);
        is($response->{margin_right}, 2);
        is($response->{alignment}, 'Center');
	};
    if ($@) {
        fail("create_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'update paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::Paragraph->new();
        $dto->{margin_left} = 2;
        $dto->{margin_right} = 2;
        $dto->{alignment} = 'Center';
        
		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 2,
            'paragraph_index' => 1,
            'dto' => $dto, 
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->update_paragraph(%params);
        
        is($response->{margin_left}, 2);
        is($response->{margin_right}, 2);
        is($response->{alignment}, 'Center');
	};
    if ($@) {
        fail("update_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'update sub-shape paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::Paragraph->new();
        $dto->{margin_left} = 2;
        $dto->{margin_right} = 2;
        $dto->{alignment} = 'Center';
        
		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 3,
            'paragraph_index' => 1,
            'dto' => $dto, 
            'password' => 'password', 
            'folder' => 'TempSlidesSDK',
            'sub_shape' => '1');
        my $response = $utils->{testSlidesApi}->update_paragraph(%params);
        
        is($response->{margin_left}, 2);
        is($response->{margin_right}, 2);
        is($response->{alignment}, 'Center');
	};
    if ($@) {
        fail("update_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'delete paragraphs' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);
        
		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 2,
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->delete_paragraphs(%params);
        
        is(scalar @{$response->{paragraph_links}}, 0);
	};
    if ($@) {
        fail("delete_paragraphs raised an exception: $@");
    }
    pass();
};

subtest 'delete paragraphs indexes' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);
        my @paragraphs = (2);
		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 2,
            'paragraphs' => \@paragraphs,
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->delete_paragraphs(%params);
        
        is(scalar @{$response->{paragraph_links}}, 1);
	};
    if ($@) {
        fail("delete_paragraphs raised an exception: $@");
    }
    pass();
};

subtest 'delete sub-shape paragraphs' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);
        
		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 3,
            'password' => "password", 
            'folder' => "TempSlidesSDK",
            'sub_shape' => '1');
        my $response = $utils->{testSlidesApi}->delete_paragraphs(%params);
        
        is(scalar @{$response->{paragraph_links}}, 0);
	};
    if ($@) {
        fail("delete_paragraphs raised an exception: $@");
    }
    pass();
};

subtest 'delete sub-shape paragraphs indexes' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);
        my @paragraphs = (1);
		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 3,
            'paragraphs' => \@paragraphs,
            'password' => "password", 
            'folder' => "TempSlidesSDK",
            'sub_shape' => '1');
        my $response = $utils->{testSlidesApi}->delete_paragraphs(%params);
        
        is(scalar @{$response->{paragraph_links}}, 1);
	};
    if ($@) {
        fail("delete_paragraphs raised an exception: $@");
    }
    pass();
};

subtest 'delete paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);
        
		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 2,
            'paragraph_index' => 1,
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->delete_paragraph(%params);
        
        is(scalar @{$response->{paragraph_links}}, 1);
	};
    if ($@) {
        fail("delete_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'delete sub-shape paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);
        
		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 3,
            'paragraph_index' => 1,
            'password' => "password", 
            'folder' => "TempSlidesSDK",
            'sub_shape' => '1');
        my $response = $utils->{testSlidesApi}->delete_paragraph(%params);
        
        is(scalar @{$response->{paragraph_links}}, 1);
	};
    if ($@) {
        fail("delete_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'get paragraph rect' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);
        
		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 2,
            'paragraph_index' => 1,
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->get_paragraph_rectangle(%params);
        
        ok($response->{x} > 0);
        ok($response->{y} > 0);
        ok($response->{width} > 0);
        ok($response->{height} > 0);
	};
    if ($@) {
        fail("get_paragraph_rectangle raised an exception: $@");
    }
    pass();
};

subtest 'paragraph default portion format' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $dto = AsposeSlidesCloud::Object::Paragraph->new();
        my $default_portion_format = AsposeSlidesCloud::Object::PortionFormat->new();
        $default_portion_format->{font_height} = 20;
        $default_portion_format->{latin_font} = 'Arial';
        $dto->{default_portion_format} = $default_portion_format;
        my $portion1 = AsposeSlidesCloud::Object::Portion->new();
        $portion1->{text} = 'Portion1';
        my $portion2 = AsposeSlidesCloud::Object::Portion->new();
        $portion2->{text} = 'Portion2';
        my @portions = ($portion1, $portion2);
        $dto->{portion_list} = \@portions;
        
		my %params = ('name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 2,
            'dto' => $dto,
            'password' => "password", 
            'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->create_paragraph(%params);

        is(scalar @{$response->{portion_list}}, 2);
        is($response->{default_portion_format}->{latin_font}, $dto->{default_portion_format}->{latin_font});
        is($response->{default_portion_format}->{font_height}, $dto->{default_portion_format}->{font_height});
	};
    if ($@) {
        fail("create_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'get paragraph effective' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = ('name' => "test.pptx", 'slide_index' => 6, 'shape_index' => 2, 'paragraph_index' => 1, 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->get_paragraph_effective(%params);
        is($response->{default_tab_size}, 72);
	};
    if ($@) {
        fail("get_paragraph_effective raised an exception: $@");
    }
    pass();
};

subtest 'get sub-shape paragraph effective' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

		my %params = (
            'name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 3, 
            'paragraph_index' => 1, 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'sub_shape' => '1');

        my $response = $utils->{testSlidesApi}->get_paragraph_effective(%params);
        is($response->{default_tab_size}, 72);
	};
    if ($@) {
        fail("get_paragraph_effective raised an exception: $@");
    }
    pass();
};

subtest 'download paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'slide_index' => 6, 'shape_index' => 2, 'paragraph_index' => 1, 'format' => 'Png', 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->download_paragraph(%params);
        ok(length($response) > 0);
    };
    if ($@) {
        fail("download_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'download sub-shape paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = (
            'name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 3, 
            'paragraph_index' => 1, 
            'format' => 'Png', 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'sub_shape' => '1');
        my $response = $utils->{testSlidesApi}->download_paragraph(%params);
        ok(length($response) > 0);
    };
    if ($@) {
        fail("download_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'download paragraph scaled' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my %params = ('name' => "test.pptx", 'slide_index' => 6, 'shape_index' => 2, 'paragraph_index' => 1, 'format' => 'Png', 'password' => "password", 'folder' => "TempSlidesSDK");
        my $response = $utils->{testSlidesApi}->download_paragraph(%params);
        $params{'scale_x'} = 2;
        $params{'scale_y'} = 2;
        my $scaled_response = $utils->{testSlidesApi}->download_paragraph(%params);
        ok(length($scaled_response) > length($response));
    };
    if ($@) {
        fail("download_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'save paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $out_path = "TempSlidesSDK/paragraph.png";
        my %params = ('name' => "test.pptx", 'slide_index' => 6, 'shape_index' => 2, 'paragraph_index' => 1, 'format' => 'Png', 'out_path' => $out_path, 'password' => "password", 'folder' => "TempSlidesSDK");
        $utils->{testSlidesApi}->save_paragraph(%params);

        my %exists_params = ('path' => $out_path);
        my $exists = $utils->{testSlidesApi}->object_exists(%exists_params);
        ok($exists->{exists});
    };
    if ($@) {
        fail("save_paragraph raised an exception: $@");
    }
    pass();
};

subtest 'save sub-shape paragraph' => sub {
    eval {
        my %copy_params = ('src_path' => "TempTests/test.pptx", 'dest_path' => "TempSlidesSDK/test.pptx");
        $utils->{testSlidesApi}->copy_file(%copy_params);

        my $out_path = "TempSlidesSDK/paragraph.png";
        my %params = (
            'name' => "test.pptx", 
            'slide_index' => 6, 
            'shape_index' => 3, 
            'paragraph_index' => 1, 
            'format' => 'Png', 
            'out_path' => $out_path, 
            'password' => "password", 
            'folder' => "TempSlidesSDK", 
            'sub_shape' => '1');
        $utils->{testSlidesApi}->save_paragraph(%params);

        my %exists_params = ('path' => $out_path);
        my $exists = $utils->{testSlidesApi}->object_exists(%exists_params);
        ok($exists->{exists});
    };
    if ($@) {
        fail("save_paragraph raised an exception: $@");
    }
    pass();
};

done_testing;