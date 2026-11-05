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

package AsposeSlidesCloud::TestUtils;

use File::Slurp;
use Test::More;
use JSON;

use strict;
use warnings;
use utf8;

use AsposeSlidesCloud::ClassRegistry;
use AsposeSlidesCloud::Configuration;

use AsposeSlidesCloud::SlidesApi;
use AsposeSlidesCloud::SlidesAsyncApi;

my $is_initialized = 0;
my $test_data_version = "1";
my $test_operation_id = "";

sub new {
    my $class = shift;
    my %params = @_;
    my $config_file = decode_json(read_file("testConfig.json"));

    my $config = AsposeSlidesCloud::Configuration->new();
    $config->{base_url} = $config_file->{BaseUrl};
    $config->{auth_base_url} = $config_file->{AuthBaseUrl};
    $config->{app_sid} = $config_file->{ClientId};
    $config->{app_key} = $config_file->{ClientSecret};
    $config->{debug} = $config_file->{Debug};
    $config->{allow_insecure_requests} = $config_file->{AllowInsecureRequests};
    my $api = AsposeSlidesCloud::SlidesApi->new(config => $config);

    my $asyncConfig = AsposeSlidesCloud::Configuration->new();
    $asyncConfig->{base_url} = $config_file->{AsyncBaseUrl};
    $asyncConfig->{auth_base_url} = $config_file->{AuthBaseUrl};
    $asyncConfig->{app_sid} = $config_file->{ClientId};
    $asyncConfig->{app_key} = $config_file->{ClientSecret};
    $asyncConfig->{debug} = $config_file->{Debug};
    $asyncConfig->{allow_insecure_requests} = $config_file->{AllowInsecureRequests};
    my $asyncApi = AsposeSlidesCloud::SlidesAsyncApi->new(config => $asyncConfig);

    return bless { rules => decode_json(read_file('testRules.json')), testSlidesApi => $api, testSlidesAsyncApi => $asyncApi }, $class;
}

sub initialize {
    my ($self, $function, $parameter, $parameter_value, $parameter_type) = @_;
    $function =~ s/_//g;
    $parameter =~ s/_//g;
    if (!$is_initialized) {
        my %download_params = ('path' => 'TempTests/version.txt');
        my $version = $self->{testSlidesApi}->download_file(%download_params);
        if ($version != $test_data_version) {
            opendir my $dir, "TestData";
            my @files = readdir $dir;
            closedir $dir;
            foreach (@files) {
                if(-f "TestData/".$_) {
                    my $content = read_file("TestData/".$_, { binmode => ':raw' });
                    my %file_upload_params = ('path' => 'TempTests/'.$_, 'file' => $content);
                    $self->{testSlidesApi}->upload_file(%file_upload_params);
                }
            }
            my %upload_params = ('path' => 'TempTests/version.txt', 'file' => $test_data_version);
            $self->{testSlidesApi}->upload_file(%upload_params);
        }
        $is_initialized = 1;
    }

    my %files = ();
    my $fileRules = $self->{rules}->{Files};
    foreach (@$fileRules) {
        if ($self->is_good_rule($_, $function, $parameter, $parameter_type)) {
            my $actual_name = $self->untemplatize($_->{File}, $parameter, $parameter_value);
            my $path = "TempSlidesSDK";
            if (exists $_->{Folder}) {
                $path = $self->untemplatize($_->{Folder}, $parameter, $parameter_value);
            }
            $path = $path."/".$actual_name;
            $files{$path} = $_;
            $_->{ActualName} = $actual_name;
        }
    }
    foreach my $path (keys %files) {
        if ($files{$path}->{Action} eq "Put") {
            my %copy_params = ('src_path' => 'TempTests/'.$files{$path}->{ActualName}, 'dest_path' => $path);
            $self->{testSlidesApi}->copy_file(%copy_params);
        } elsif ($files{$path}->{Action} eq "Delete") {
            my %delete_params = ('path' => $path);
            $self->{testSlidesApi}->delete_file(%delete_params);
            $self->{testSlidesApi}->delete_folder(%delete_params);
        }
    }
}

# Set the user agent of the API client
#
# @param string $user_agent The user agent of the API client
#
sub get_param_value {
    my ($self, $function, $parameter, $type) = @_;
    $function =~ s/_//g;
    $parameter =~ s/_//g;
    my $result = undef;
    my $values = $self->{rules}->{Values};
    foreach (@$values) {
        if ($self->is_good_rule($_, $function, $parameter, $type)) {
            if (exists $_->{Value}) {
                $result = $_->{Value}
            }
        }
    }
    return $self->untemplatize($result, $parameter, "");
}

sub invalidize_param_value {
    my ($self, $function, $parameter, $value, $type) = @_;
    $function =~ s/_//g;
    $parameter =~ s/_//g;
    my $result = undef;
    my $values = $self->{rules}->{Values};
    foreach (@$values) {
        if ($self->is_good_rule($_, $function, $parameter, $type)) {
            if (exists $_->{InvalidValue}) {
                $result = $_->{InvalidValue};
            }
        }
    }
    if (!$result) {
        return $result;
    }
    return $self->untemplatize($result, $parameter, $value);
}

sub assert_error {
    my ($self, $function, $parameter, $value, $type, $error) = @_;
    if ($error) {
        my $code = 0;
        my $message = "unexpected message";
        my $results = $self->{rules}->{Results};
        $function =~ s/_//g;
        $parameter =~ s/_//g;
        foreach (@$results) {
            if ($self->is_good_rule($_, $function, $parameter, $type)) {
                if (exists $_->{Code}) {
                    $code = $_->{Code};
                }
                if (exists $_->{Message}) {
                    $message = $_->{Message};
                }
            }
        }
        if ($error =~ m/API Exception\((\d+)\): (.*) /s) {
            is($1, $code);
        } else {
            is(400, $code);
        }
        ok(index($error, $self->untemplatize($message, $parameter, $value)) != -1);
    } else {
        fail("expected to fail for $function $parameter");
    }
}

sub assert_no_error {
    my ($self, $function, $parameter, $type) = @_;
    my $results = $self->{rules}->{OKToNotFail};
    $function =~ s/_//g;
    $parameter =~ s/_//g;
    foreach (@$results) {
        if ($self->is_good_rule($_, $function, $parameter, $type)) {
            pass();
            return;
        }
    }
    fail("expected to fail for $function $parameter");
}

sub is_good_rule {
    my ($self, $rule, $function, $parameter, $type) = @_;
    if ($self->is_good_value($rule->{Method}, $function)
        and (!(defined $rule->{Invalid}) or ($rule->{Invalid} eq defined $parameter))
        and $self->is_good_value($rule->{Parameter}, $parameter)
        and $self->is_good_type($rule, $type)
        and $self->is_good_value($rule->{Language}, "PERL")) {
        return 1;
    } else {
        return 0;
    };
}

sub is_good_value {
    my ($self, $rule_value, $value) = @_;
    if (!(defined $rule_value)) {
        return 1;
    }
    if (!$value) {
        return 0;
    }
    if ('/' eq substr($rule_value, 0, 1) and '/' eq substr($rule_value, -1)) {
        my $regex = substr($rule_value, 1, -1);
        return $value =~ /$regex/i;
    }
    return uc($rule_value) eq uc($value);
}

sub is_good_type {
    my ($self, $rule, $type) = @_;
    if (!(defined $rule->{Type})) {
        return 1;
    }
    if (!$type) {
        return 0;
    }
    if ($rule->{Type} eq 'bool') {
        return $type eq 'boolean';
    }
    if ($rule->{Type} eq 'number') {
        return ($type eq 'int') || ($type eq 'double');
    }
    if ($rule->{Type} eq 'int') {
        return $type eq 'int';
    }
    if ($rule->{Type} eq 'int[]') {
        return $type eq 'int[]';
    }
    if ($rule->{Type} eq 'string[]') {
        return $type eq 'string[]';
    }
    if ($rule->{Type} eq 'stream') {
        return $type eq 'File';
    }
    if ($rule->{Type} eq 'stream[]') {
        return $type eq 'ARRAY[string]';
    }
    if ($rule->{Type} eq 'model') {
        return AsposeSlidesCloud::ClassRegistry->has_class($type);
    }
    if (AsposeSlidesCloud::ClassRegistry->has_class($rule->{Type})) {
        return AsposeSlidesCloud::ClassRegistry->is_subclass($rule->{Type}, $type);
    }
    return 0;
}

sub untemplatize {
    my ($self, $template, $name, $value) = @_;
    if (!$template) {
        return $value;
    } else {
        if ('@' eq substr($template, 0, 1)) {
            my $file_name = substr($template, 1);
            if ('(' eq substr($file_name, 0, 1) and ')' eq substr($file_name, -1)) {
                my @file_names = split(",", substr($file_name, 1, -1));
                my @files = ();
                foreach (@file_names) {
                    my $file = read_file("TestData/$_", { binmode => ':raw' });
                    push @files, $file;
                }
                return \@files;
            }
            my $content = read_file("TestData/$file_name", { binmode => ':raw' });
            return $content;
        }
        if ('#OperationId' eq $template) {
            return $self->get_operation_id();
        }
        if ('#NewId' eq $template) {
            return "541f91c7-5721-4f0f-9bc0-0d18c1b4ae82";
        }
        my $result = $template;
        if (defined $value) {
            $result =~ s/%v/$value/;
        } else {
            $result =~ s/%v//;
        }
        if (defined $name) {
            $result =~ s/%n/$name/;
        } else {
            $result =~ s/%n//;
        }
        return $result;
    }
}

sub get_operation_id {
    my ($self) = @_;
    if ('' eq $test_operation_id) {
        my $source = read_file("TestData/test.pptx", { binmode => ':raw' });
        my %params = ('document' => $source, 'format' => 'pdf', 'password' => 'password');
        $test_operation_id = $self->{testSlidesAsyncApi}->start_convert(%params);
        sleep(10);
    }
    return $test_operation_id;
}

1;
