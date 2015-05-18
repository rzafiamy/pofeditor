#-----------------------------------------------------------------------------
#
#	$Id : XPath.pm 2.237 2010-07-12 JMG$
#
#	Created and maintained by Jean-Marie Gouarne
#	Copyright 2010 by Genicorp, S.A. (www.genicorp.com)
#
#-----------------------------------------------------------------------------

package	OpenOffice::OODoc::XPath;
use	5.008_000;
use     strict;
our	$VERSION	= '2.237';
use	XML::Twig	3.32;
use	Encode;
require	Exporter;
our	@ISA	= qw    ( Exporter );
our	@EXPORT	= qw
                        (
                        TRUE FALSE is_true is_false
                        odfLocaltime odfTimelocal
                        );

#------------------------------------------------------------------------------

use constant
        {
        TRUE    => 1,
        FALSE   => 0
        };

sub     is_true
        {
        my $arg = shift         or return FALSE;
        $arg    = lc $arg;
        return ($arg eq '1' || $arg eq 'true' || $arg eq 'on') ? TRUE : FALSE;
        }

sub     is_not_true
        {
        return is_true(shift) ? FALSE : TRUE; 
        }

#------------------------------------------------------------------------------

BEGIN	{
	*dispose		= *DESTROY;
	*update			= *save;
	*getXMLContent		= *exportXMLContent;
	*getContent		= *exportXMLContent;
	*getChildElementByName	= *selectChildElementByName;
	*getElementByIdentifier = *selectElementByIdentifier;
	*blankSpaces		= *spaces;
	*createSpaces		= *spaces;
	*createTextNode         = *newTextNode;
	*getFrame		= *getFrameElement;
	*getUserFieldElement	= *getUserField;
	*getVariableElement     = *getVariable;
	*getNodeByXPath		= *selectNodeByXPath;
	*getNodesByXPath	= *selectNodesByXPath;
	*getElementList         = *selectNodesByXPath;
	*isCalcDocument		= *isSpreadsheet;
	*isDrawDocument		= *isDrawing;
	*isImpressDocument	= *isPresentation;
	*isWriterDocument	= *isText;
	*odfVersion		= *openDocumentVersion;
	}

#------------------------------------------------------------------------------

our %XMLNAMES	=			# OODoc root element names
	(
	'content'	=> 'office:document-content',
	'styles'	=> 'office:document-styles',
	'meta'		=> 'office:document-meta',
	'manifest'	=> 'manifest:manifest',
	'settings'	=> 'office:document-settings'
	);

					# characters to be escaped in XML
our	$CHARS_TO_ESCAPE	= "\"<>'&";
					# standard external character set
our	$LOCAL_CHARSET		= 'utf8';
					# standard ODF character set
our	$OO_CHARSET		= 'utf8';
                                        # default element identifier
our     $ELT_ID                 = 'text:id';

#------------------------------------------------------------------------------
# basic conversion between internal & printable encodings

sub	OpenOffice::OODoc::XPath::decode_text
	{
	return Encode::encode($LOCAL_CHARSET, shift);
	}

sub	OpenOffice::OODoc::XPath::encode_text
	{
	return Encode::decode($LOCAL_CHARSET, shift);
	}

#------------------------------------------------------------------------------
# common date formatting functions

sub	odfLocaltime
	{
	my $time = shift || time();
	my @t = localtime($time);
	return sprintf
			(
			"%04d-%02d-%02dT%02d:%02d:%02d",
			$t[5] + 1900, $t[4] + 1, $t[3], $t[2], $t[1], $t[0]
			);
	}

sub	odfTimelocal
	{
	require Time::Local;

	my $ootime = shift;
	return undef unless $ootime;
	$ootime =~ /(\d*)-(\d*)-(\d*)T(\d*):(\d*):(\d*)/;
	return Time::Local::timelocal($6, $5, $4, $3, $2 - 1, $1); 
	}

#------------------------------------------------------------------------------
# object coordinates, size, description control

sub	setObjectCoordinates
	{
	my $self	= shift;
	my $element	= shift	or return undef;
	my ($x, $y)	= @_;
	if ($x && ($x =~ /,/))	# X and Y are concatenated in a single string
		{
		$x =~ s/\s*//g;			# remove the spaces
		$x =~ s/,(.*)//; $y = $1;	# split on the comma
		}
	$x = '0cm' unless $x; $y = '0cm' unless $y;
	$x .= 'cm' unless $x =~ /[a-zA-Z]$/;
	$y .= 'cm' unless $y =~ /[a-zA-Z]$/;
	$self->setAttributes($element, 'svg:x' => $x, 'svg:y' => $y);
	return wantarray ? ($x, $y) : ($x . ',' . $y);
	}

sub	getObjectCoordinates
	{
	my $self	= shift;
	my $element	= shift	or return undef;
	my $x		= $element->getAttribute('svg:x');
	my $y		= $element->getAttribute('svg:y');
	return undef unless defined $x and defined $y;
	return wantarray ? ($x, $y) : ($x . ',' . $y);
	}

sub	setObjectSize
	{
	my $self	= shift;
	my $element	= shift	or return undef;
	my ($w, $h)	= @_;
	if ($w && ($w =~ /,/))	# W and H are concatenated in a single string
		{
		$w =~ s/\s*//g;			# remove the spaces
		$w =~ s/,(.*)//; $h = $1;	# split on the comma
		}
	$w = '0cm' unless $w; $h = '0cm' unless $h;
	$w .= 'cm' unless $w =~ /[a-zA-Z]$/;
	$h .= 'cm' unless $h =~ /[a-zA-Z]$/;
	$self->setAttributes($element, 'svg:width' => $w, 'svg:height' => $h);
	return wantarray ? ($w, $h) : ($w . ',' . $h);
	}

sub	getObjectSize
	{
	my $self	= shift;
	my $element	= shift	or return undef;
	my $w		= $element->getAttribute('svg:width');
	my $h		= $element->getAttribute('svg:height');
	return wantarray ? ($w, $h) : ($w . ',' . $h);
	}

sub	setObjectDescription
	{
	my $self	= shift;
	my $element	= shift or return undef;
	my $text	= shift;
	my $desc	= $element->first_child('svg:desc');
	unless ($desc)
		{
		$self->appendElement($element, 'svg:desc', text => $text)
			if (defined $text);
		}
	else
		{
		if (defined $text)	{ $self->setText($desc, $text, @_);	}
		else			{ $self->removeElement($desc, @_);	}
		}
	return $desc;
	}

sub	getObjectDescription
	{
	my $self	= shift;
	my $element	= shift or return undef;
	return $self->getXPathValue($element, 'svg:desc');
	}

sub     getObjectName
        {
	my $self	= shift;
	my $element	= shift or return undef;
	my $name	= shift;
	my $attr        = $element->getPrefix() . ':name' ;
        return $self->getAttribute($element, $attr);        
        }

sub     setObjectName
        {
	my $self	= shift;
	my $element	= shift or return undef;
	my $name	= shift;
	my $attr        = $element->getPrefix() . ':name' ;
        return $self->setAttribute($element, $attr, @_);        
        }

sub	objectName
	{
	my $self	= shift;
	my $element	= shift or return undef;
	my $name	= shift;
	my $attr        = $element->getPrefix() . ':name' ;
	return (defined $name) ?
		$self->setAttribute($element, $attr => $name)	:
		$self->getAttribute($element, $attr);
	}

#------------------------------------------------------------------------------
# basic element creation

sub	OpenOffice::OODoc::XPath::new_element
	{
	my $name	= shift		or return undef;
	return undef if ref $name;
	$name		=~ s/^\s+//;
	$name		=~ s/\s+$//;
	if ($name =~ /^</)	# create element from XML string
		{
		return OpenOffice::OODoc::Element->parse($name, @_);
		}
	else			# create element from name and optional data
		{
		return OpenOffice::OODoc::Element->new($name, @_);
		}
	}

#------------------------------------------------------------------------------
# text node creation

sub	OpenOffice::OODoc::XPath::new_text_node
	{
	return OpenOffice::OODoc::XPath::new_element('#PCDATA', @_);
	}

#------------------------------------------------------------------------------
# basic conversion between internal & printable encodings (object version)

sub	inputTextConversion
	{
	my $self	= shift;
	my $text	= shift;
	return undef unless defined $text;
	my $local_encoding = $self->{'local_encoding'} or return $text;
	return Encode::decode($local_encoding, $text);
	}

sub	outputTextConversion
	{
	my $self	= shift;
	my $text	= shift;
	return undef unless defined $text;
	my $local_encoding = $self->{'local_encoding'} or return $text;
	return Encode::encode($local_encoding, $text);
	}

sub	localEncoding
	{
	my $self	= shift;
	my $encoding	= shift;
	$self->{'local_encoding'} = $encoding if $encoding;
	return $self->{'local_encoding'} || '';
	}

sub	noLocalEncoding
	{
	my $self	= shift;
	delete $self->{'local_encoding'};
	return 1;
	}

#------------------------------------------------------------------------------
# search/replace text processing routine
# if $replace is a user-provided routine, it's called back with
# the current argument stack, plus the substring found

sub	_find_text
	{
	my $self	= shift;
	my $text	= shift;
	my $pattern	= $self->inputTextConversion(shift);
	my $replace	= shift;

	if (defined $pattern)
	    {
	    if (defined $replace)
		{
		if (ref $replace)
		    {
		    if ((ref $replace) eq 'CODE')
		    	{
			return undef
			  unless
			    (
			    $text =~
			    	s/($pattern)/
				    	{
					my $found = $1;
					Encode::_utf8_on($found)
						if Encode::is_utf8($text);
					my $result = &$replace(@_, $found);
					$result = $found
						unless (defined $result);
					$result;
					}
				/eg
			    );
			}
		    else
		    	{
			return undef unless ($text =~ /$pattern/);
			}
		    }
		else
		    {
		    my $r = $self->inputTextConversion($replace);
		    return undef unless ($text =~ s/$pattern/$r/g);
		    }
		}
	    else
		{
		return undef unless ($text =~ /$pattern/);
		}
	    }
	return $text;
	}

#------------------------------------------------------------------------------
# search/replace content in descendant nodes

sub	_search_content
	{
	my $self	= shift;
	my $node	= shift or return undef;
	my $content	= undef;

        if ($node->isTextNode)
                {
                my $text = $self->_find_text($node->text, @_);
                if (defined $text)
                        {
                        $node->set_text($text);
                        $content = $text;
                        }
                }    
        else
                {        
	        foreach my $n ($node->getTextDescendants)
		        {
		        my $text = $self->_find_text($n->text, @_);
		        if (defined $text)
			        {
			        $n->set_text($text);
			        $content .= $text;
			        }
			}
		}
	return $content;
	}
	
#------------------------------------------------------------------------------
# is this an OASIS Open Document or an OpenOffice 1.x Document ?

sub	isOpenDocument
	{
	my $self	= shift;
	my $root	= $self->getRootElement;
	die __PACKAGE__ . " Missing root element\n" unless $root;
	my $ns		= $root->att('xmlns:office');
	return $ns && ($ns =~ /opendocument/) ? 1 : undef;
	}

sub	openDocumentVersion
	{
	my $self	= shift;
	my $new_version	= shift;
	my $root	= $self->getRootElement or return undef;
	$root->set_att('office:version' => $new_version) if $new_version;
	return $root->att('office:version');
	}

#------------------------------------------------------------------------------
# document class check

sub	isContent
	{
	my $self	= shift;
	return ($self->contentClass()) ? 1 : undef;
	}

sub	isSpreadsheet
	{
	my $self	= shift;
	return ($self->contentClass() eq 'spreadsheet') ? 1 : undef;
	}
sub	isPresentation
	{
	my $self	= shift;
	return ($self->contentClass() eq 'presentation') ? 1 : undef;
	}
sub	isDrawing
	{
	my $self	= shift;
	return ($self->contentClass() eq 'drawing') ? 1 : undef;
	}
sub	isText
	{
	my $self	= shift;
	return ($self->contentClass() eq 'text') ? 1 : undef;
	}

#------------------------------------------------------------------------------

sub     _get_container      # get a new OODoc::File container
        {
        require OpenOffice::OODoc::File;
        
        my $doc         = shift;
                
	return OpenOffice::OODoc::File->new
				(
				$doc->{'file'},
				create		=> $doc->{'create'},
				opendocument	=> $doc->{'opendocument'},
				template_path	=> $doc->{'template_path'}
				);
        }
        
sub     _get_flat_file          # get flat ODF content
        {
        my $doc         = shift;
        my $source      = $doc->{'file'};
	$doc->{'xpath'} = UNIVERSAL::isa($source, 'IO::File') ?
			     $doc->{'twig'}->safe_parse($source)    :
			     $doc->{'twig'}->safe_parsefile($source);
        return $doc->{'path'};
        }

sub	new
	{
	my $caller	= shift;
	my $class	= ref($caller) || $caller;
	my $self	=
		{
		auto_style_path		=> '//office:automatic-styles',
		master_style_path	=> '//office:master-styles',
		named_style_path	=> '//office:styles',
		image_container		=> 'draw:image',
		image_xpath		=> '//draw:image',
		image_fpath		=> '#Pictures/',
		local_encoding		=>
				$OpenOffice::OODoc::XPath::LOCAL_CHARSET,
		@_
		};
	
	foreach my $optk (keys %$self)
		{
		next unless $self->{$optk};
		my $v = lc $self->{$optk};
		$self->{$optk} = 0 if ($v =~ /^false$|^off$/);
		}

	$self->{'container'} = $self->{'file'} if defined $self->{'file'};
	$self->{'container'} = $self->{'archive'} if defined $self->{'archive'};
	$self->{'part'} = $self->{'member'} if $self->{'member'};
	$self->{'part'} = 'content' unless $self->{'part'};

	unless ($self->{'element'})
		{
		my $m	= lc $self->{'part'};
		if ($m =~ /(^.*)\..*/) { $m = $1; }
		$self->{'element'} =
		    $OpenOffice::OODoc::XPath::XMLNAMES{$m};
		}
					# create the XML::Twig
	if 	(is_true($self->{'readable_XML'}))
			{
			$self->{'readable_XML'} = 'indented';
			}
	$self->{'element'} = $OpenOffice::OODoc::XPath::XMLNAMES{'content'}
				unless $self->{'element'};
	if ($self->{'element'})
		{
		$self->{'twig'} = XML::Twig->new
			(
			elt_class	=> "OpenOffice::OODoc::Element",
			twig_roots	=>
				{
				$self->{'element'}	=> 1
				},
			pretty_print	=> $self->{'readable_XML'},
			%{$self->{'twig_options'}}
			);
		}
	else
		{
		$self->{'twig'} = XML::Twig->new
			(
			elt_class	=> "OpenOffice::OODoc::Element",
			pretty_print	=> $self->{'readable_XML'},
			%{$self->{'twig_options'}}
			);
		}

	                                        # other OODoc::Xpath object
	$self->{'container'} = $self->{'container'}->{'container'}
	        if      (
	                ref($self->{container})
	                        &&
	                $self->{'container'}->isa('OpenOffice::OODoc::XPath')
                        );
	
	if ($self->{'xml'})			# load from XML string
		{
		delete $self->{'container'};
		delete $self->{'file'};
		$self->{'xpath'} =
			$self->{'twig'}->safe_parse($self->{'xml'});
		delete $self->{'xml'};
		}	
	
	elsif (defined $self->{'container'})
		{
		delete $self->{'file'};
	 	                                # existing OODoc::File object
	 	if 
	 	        (
	 	        UNIVERSAL::isa($self->{'container'},
	 	        'OpenOffice::OODoc::File')
	 	        )
	 	        {
	 	        my $xml = $self->{'container'}->link($self);
	 	        $self->{'xpath'} = $self->{'twig'}->safe_parse($xml);
	 	        }
	 	                                # source file or filehandle
	 	else
	 	        {
	 	        $self->{'file'} = $self->{'container'};
	 	        delete $self->{'container'};
	 	        if	(
	 	                $self->{'flat_xml'}
					||
			        (lc $self->{'file'}) =~ /\.xml$/
			        )
			        		# XML flat file
			        {
			        $self->{'xpath'} = _get_flat_file($self);
			        }
		        else
			        {		# new OODoc::File object
			        $self->{'container'} = _get_container($self);
			        return undef unless $self->{'container'};
			        delete $self->{'file'};
			        my $xml = $self->{'container'}->link($self);
			        $self->{'xpath'} =
			                $self->{'twig'}->safe_parse($xml);
			        }	
	 	        }				
		}

	unless ($self->{'xpath'})
		{
		warn "[" . __PACKAGE__ . "::new] No ODF content\n";
		return undef;
		}
						# XML content loaded & parsed
	bless $self, $class;
	
	$self->{'opendocument'} = $self->isOpenDocument;
	
	if ($self->{'opendocument'})
		{
		$self->{'image_container'}	= 'draw:frame';
		$self->{'image_xpath'}		= '//draw:frame';
		$self->{'image_fpath'}		= 'Pictures/';
		}
	
	$self->{'member'} = $self->{'part'};		# for compatibility
	$self->{'archive'} = $self->{'container'};	# for compatibility
	$self->{'context'} = $self->getRoot;
	$self->{'body'} = $self->getBody;

	return $self;
	}

#------------------------------------------------------------------------------
# destructor

sub	DESTROY
	{
	my $self	= shift;

	if ($self->{'body'})
		{
		$self->{'body'}->dispose();
		}
	delete $self->{'body'};
	if ($self->{'context'})
	        {
	        $self->{'context'}->dispose();
	        }
	delete $self->{'context'};
	if ($self->{'xpath'})
		{
		$self->{'xpath'}->dispose();
		}
	delete $self->{'xpath'};
	if ($self->{'twig'})
		{
		$self->{'twig'}->dispose();
		}
	delete $self->{'twig'};
	delete $self->{'xml'};
	delete $self->{'content_class'};
	delete $self->{'file'};
	delete $self->{'container'};
	delete $self->{'archive'};
	delete $self->{'part'};
	delete $self->{'twig_options'};
	$self = {};
	}

#------------------------------------------------------------------------------
# get a reference to the embedded XML parser for share

sub	getXMLParser
	{
	warn	"[" . __PACKAGE__ . "::getXMLParser] No longer implemented\n";
	return undef;
	}

#------------------------------------------------------------------------------
# make the changes persistent in an OpenOffice.org file

sub	save
	{
	my $self	= shift;
	my $target	= shift;

	my $filename	= ($target) ? $target : $self->{'file'};
	my $archive	= $self->{'container'};
	unless ($archive)
		{
		return undef if is_true($self->{'read_only'});

		if ($filename)
			{
			open my $fh, ">:utf8", $filename;
			$self->exportXMLContent($fh);
			close $fh;
			return $filename;
			}
		else
			{
			warn "[" . __PACKAGE__ . "::save] Missing file\n";
			return undef;
			}
		}
	$filename	= $archive->{'source_file'}	unless $filename;
	unless ($filename)
		{
		warn "[" . __PACKAGE__ . "::save] No target file\n";
		return undef;
		}

	unless ($self->{'part'})
		{
		warn "[" . __PACKAGE__ . "::save] Missing archive part name\n";
		return undef;
		}

	my $result = $archive->save($filename);
	return $result;
	}

#------------------------------------------------------------------------------
# raw file import

sub	raw_import
	{
	my $self	= shift;
	if ($self->{'container'})
		{
		my $target	= shift;
		unless ($target)
			{
			warn	"[" . __PACKAGE__ . "::raw_import] "	.
				"No target member for import\n";
			return undef;
			}
		$target =~ s/^#//;
		return $self->{'container'}->raw_import($target, @_);
		}
	else
		{
		warn	"[" . __PACKAGE__ . "::raw_import] "	.
			"No container for file import\n";
		return undef;
		}
	}

#------------------------------------------------------------------------------
# raw file export

sub	raw_export
	{
	my $self	= shift;
	if ($self->{'container'})
		{
		my $source	= shift;
		unless ($source)
			{
			warn	"[" . __PACKAGE__ . "::raw_import] "	.
				"Missing source file name\n";
			return undef;
			}
		$source =~ s/^#//;
		return $self->{'container'}->raw_export($source, @_);
		}
	else
		{
		warn	"[" . __PACKAGE__ . "::raw_import] "	.
			"No container for file export\n";
		return undef;
		}
	}

#------------------------------------------------------------------------------
# exports the whole content of the document as an XML string

sub	exportXMLContent
	{
	my $self	= shift;
	my $target	= shift;
	if ($target)
		{
		return $self->{'twig'}->print($target);
		}
	else
		{
		return $self->{'twig'}->sprint;
		}
	}

#------------------------------------------------------------------------------
# brute force tree reorganization

sub	reorganize
	{
	warn "[" . __PACKAGE__ . "::reorganize] No longer implemented\n";
	return undef;
	}

#------------------------------------------------------------------------------
# returns the root of the XML document

sub	getRoot
	{
	my $self	= shift;
	return $self->{'xpath'}->root;
	}

#------------------------------------------------------------------------------
# returns the name of the document part (content, styles, meta, ...)

sub     getPartName
        {
        my $self        = shift;
        my $name        = $self->getRoot->getName;
        $name           =~ s/^office:document-//;
        return $name;
        }

#------------------------------------------------------------------------------
# returns the root element of the XML document

sub	getRootElement
	{
	my $self	= shift;

	my $root	= $self->{'xpath'}->root;
	my $rootname	= $root->name() || '';
	return ($rootname eq $self->{'element'})	?
			$root				:
			$root->first_child($self->{'element'});
	}

#------------------------------------------------------------------------------
# get/set/reset the current search context

sub	currentContext
	{
	my $self	= shift;
	my $new_context	= shift;
	$self->{'context'} = $new_context if (ref $new_context);
	return $self->{'context'};
	}

sub	resetCurrentContext
	{
	my $self	= shift;
	return $self->currentContext($self->getRoot);
	}

#------------------------------------------------------------------------------
# returns the content class (text, spreadsheet, presentation, drawing)

sub	contentClass
	{
	my $self	= shift;

	my $content_class	=
		$self->getRootElement->getAttribute('office:class');
	return $content_class if $content_class;

	my $body = $self->getBody	or return undef;
	my $name = $body->name		or return undef;
	$name =~ /(.*):(.*)/;
	return $2;
	}

#------------------------------------------------------------------------------
# element name check

sub	getRootName
	{
	my $self	= shift;
	return $self->getRootElement->name;
	}

#------------------------------------------------------------------------------
# XML part type checks

sub	isMeta
	{
	my $self	= shift;
	return ($self->getRootName() eq $XMLNAMES{'meta'}) ? 1 : undef;
	}

sub	isStyles
	{
	my $self	= shift;
	return ($self->getRootName() eq $XMLNAMES{'styles'}) ? 1 : undef;
	}

sub	isSettings
	{
	my $self	= shift;
	return ($self->getRootName() eq $XMLNAMES{'settings'}) ? 1 : undef;
	}

#------------------------------------------------------------------------------
# returns the document body element (if defined)

sub	getBody
	{
	my $self	= shift;

	return $self->{'body'} if ref $self->{'body'};
	
	my $root = $self->getRoot;
	if ($self->{'body_path'})
		{
		$self->{'body'} = $self->getElement
		                        ($self->{'body_path'}, 0, $root);
		return $self->{'body'};
		}

	my $office_body = $self->getElement('//office:body', 0, $root);
	
	if ($office_body)
		{
		$self->{'body'} = $self->{'opendocument'} ?
		    $office_body->selectChildElement
			('office:(text|spreadsheet|presentation|drawing)')
			:
		    $office_body;
		}
	else
		{
		$self->{'body'} = $self->getRootElement->selectChildElement
				(
				'office:(body|meta|master-styles|settings)'
				);
		}
		
	return $self->{'body'};
	}

#------------------------------------------------------------------------------
# makes the current OODoc::XPath object share the same content as another one

sub	cloneContent
	{
	my $self        = shift;
	my $source	= shift;

	unless ($source && $source->{'xpath'})
		{
		warn "[" . __PACKAGE__ . "::cloneContent] No valid source\n";
		return undef;
		}

	$self->{'xpath'}	= $source->{'xpath'};
	$self->{'begin'}	= $source->{'begin'};
	$self->{'xml'}		= $source->{'xml'};
	$self->{'end'}		= $source->{'end'};

	return $self->getRoot;
	}

#------------------------------------------------------------------------------
# exports an individual element as an XML string

sub	exportXMLElement
	{
	my $self	= shift;
	my $path	= shift;
	my $element	=
		(ref $path) ? $path : $self->getElement($path, shift);
	unless (defined $element)
	        {
	        warn    "[" . __PACKAGE__ . "::exportXMLElement]] "     .
	                "Missing element\n";
	        return undef;
	        }
	return $element->sprint(@_);
	}

#------------------------------------------------------------------------------
# exports the document body (if defined) as an XML string

sub	exportXMLBody
	{
	my $self	= shift;

	return	$self->exportXMLElement($self->getBody, @_);
	}

#------------------------------------------------------------------------------
# gets the reference of an XML element identified by path & position
# for subsequent processing

sub	getElement
	{
	my $self	= shift;
	my $path	= shift;
	return undef	unless $path;
	if (ref $path)
		{
		return	$path->isElementNode ? $path : undef;
		}
	my $pos		= shift || 0;
	my $context	= shift || $self->{'context'} || $self->getRoot;
	if (defined $pos && (($pos =~ /^\d*$/) || ($pos =~ /^[\d+-]\d+$/)))
		{
		my $node = $self->selectNodeByXPath($context, $path, $pos);
		return	$node && $node->isElementNode ? $node : undef;
		}
	else
		{
		warn	"[" . __PACKAGE__ . "::getElement] "	.
			"Invalid node position\n";
		return undef;
		}
	}

#------------------------------------------------------------------------------
# get the list of children (or the first child unless wantarray) matching
# a given element name and belonging to a given element

sub	selectChildElementsByName
	{
	my $self	= shift;
	my $path	= shift;
	my $element	= ref $path ? $path : $self->getElement($path, shift);
	return undef	unless $element;

	return $element->selectChildElements(@_);
	}

#------------------------------------------------------------------------------
# get the first child belonging to a given element and matching a given name

sub	selectChildElementByName
	{
	my $self	= shift;
	my $path	= shift;
	my $element	= ref $path ? $path : $self->getElement($path, shift);
	return undef			unless $element;
	return $element->selectChildElement(@_);
	}

#-----------------------------------------------------------------------------
# create a user field

sub     setUserFieldDeclaration
        {
        my $self        = shift;
        my $name        = shift         or return undef;
        my %attr        =
                        (
                        type    => 'string',
                        value   => "",
                        @_
                        );

        return undef if $self->getUserField($name);

        my $body        = $self->getBody;
        my $context     = $body->first_child('text:user-field-decls');
        unless ($context)
                {
                $context = $self->appendElement
                        ($body, 'text:user-field-decls');
                }

        
        my $va =
            (
                ($attr{'type'} eq 'float')      ||
                ($attr{'type'} eq 'currency')   ||
                ($attr{'type'} eq 'percentage')
            ) ?
                'office:value' : "office:$attr{'type'}-value" ;
        $attr{'office:value-type'}      = $attr{'type'};
        $attr{$va}                      = $attr{'value'};
        $attr{'text:name'}              = $name;
        $attr{'office:currency'}        = $attr{'currency'};
        delete @attr{qw(type value currency)};

        return $self->appendElement
                (
                $context, 'text:user-field-decl',
                attributes => { %attr }
                );
        }

#-----------------------------------------------------------------------------
# get user field element

sub	getUserField
	{
	my $self	= shift;
	my $name	= shift;

	unless ($name)
		{
		warn "[" . __PACKAGE__ . "::getUserField] Missing name\n";
		return undef;
		}
	if (ref $name)
		{
		my $n = $name->getName;
		return ($n eq 'text:user-field-decl') ? $name : undef;
		}
	$name = $self->inputTextConversion($name);
	my $context     = $self->getRoot();
	if ($self->getPartName() eq 'styles')
	        {
	        $context = shift || $self->currentContext;
	        }
	return $self->getNodeByXPath
			(
			"//text:user-field-decl[\@text:name=\"$name\"]",
			$context
			);
	}

#-----------------------------------------------------------------------------
# get user field list

sub     getUserFields
        {
        my $self        = shift;
        my $context     = $self->getRoot;

        if ($self->getPartName() eq 'styles')
                {
                $context = shift || $self->currentContext;
                }

        return $self->selectNodesByXPath('//text:user-field-decl', $context);
        }

#-----------------------------------------------------------------------------
# get/set user field value

sub	userFieldValue
	{
	my $self	= shift;
	my $field	= $self->getUserField(shift) or return undef;
	my $value	= shift;

	my $value_att	= $self->fieldValueAttributeName($field);

	if (defined $value)
		{
		if ($value)
			{
			$self->setAttribute($field, $value_att, $value);
			}
		else
			{
			$field->set_att($value_att => $value);
			}
		}
	return $self->getAttribute($field, $value_att);
	}

#-----------------------------------------------------------------------------
# get a variable element (contributed by Andrew Layton)

sub	getVariable
	{
	my $self	= shift;
	my $name	= shift;

	unless ($name) {
		warn	"[" . __PACKAGE__ . "::getVariable] " .
			"Missing name\n";
		return undef;
		}

	if (ref $name) {
		my $n = $name->getName;
		return ($n eq 'text:variable-set') ? $name : undef;
	}

	$name = $self->inputTextConversion($name);
	return $self->getNodeByXPath
	        ("//text:variable-set[\@text:name=\"$name\"]");
	}

#-----------------------------------------------------------------------------
# get/set the content of a variable element (contributed by Andrew Layton)

sub	variableValue
	{
	my $self	= shift;
	my $variable	= $self->getVariable(shift) or return undef;
	my $value	= shift;

	my $value_att	= $self->fieldValueAttributeName($variable);

	if (defined $value)
		{
		$self->setAttribute($variable, $value_att, $value);
		$self->setText($variable, $value);
		}

	$value = $self->getAttribute($variable, $value_att);
	return defined $value ? $value : $self->getText($variable);
	}

#-----------------------------------------------------------------------------
# some usual text field constructors

sub	create_field
	{
	my $self	= shift;
	my $tag		= shift;
	my %opt		= @_;
	my $prefix	= $opt{'-prefix'};
	delete $opt{'-prefix'};

	if ($prefix)
		{
		$tag = "$prefix:$tag" unless $tag =~ /:/;
		my %att = ();
		foreach my $k (keys %opt)
			{
			my $a = ($k =~ /:/) ? $k : "$prefix:$k";
			$att{$a} = $opt{$k};
			}
		%opt = %att;
		}
	my $element = OpenOffice::OODoc::Element->new($tag);
	$self->setAttributes($element, %opt);
	return $element;
	}

sub	spaces
	{
	my $self	= shift;
	my $length	= shift;
	return $self->create_field('text:s', 'text:c' => $length, @_);
	}

sub	tabStop
	{
	my $self	= shift;
	my $tag = $self->{'opendocument'} ? 'text:tab' : 'text:tab-stop';
	return $self->create_field($tag, @_);
	}

sub	lineBreak
	{
	my $self	= shift;
	return $self->create_field('text:line-break', @_);
	}

#------------------------------------------------------------------------------

sub	appendLineBreak
	{
	my $self	= shift;
	my $element	= shift;

	return $element->appendChild('text:line-break');
	}

#------------------------------------------------------------------------------

sub	appendSpaces
	{
	my $self	= shift;
	my $element	= shift;
	my $length	= shift;

	my $spaces	= $self->spaces($length) or return undef;
	$spaces->paste_last_child($element);
	}

#------------------------------------------------------------------------------
# multiple whitespace handling routine, contributed by J David Eisenberg 

sub processSpaces
	{
	my $self = shift;
	my $element = shift;
	my $str = shift;
	my @words = split(/(\s\s+)/, $str);
	foreach my $word (@words)
		{
		if ($word =~ m/^ +$/)
			{
			$self->appendSpaces($element, length($word));
			}
		elsif (length($word) > 0)
			{
			$element->appendTextChild($word);
			}
		}
	}

#------------------------------------------------------------------------------

sub	appendTabStop
	{
	my $self	= shift;
	my $element	= shift;

	my $tabtag = $self->{'opendocument'} ? 'text:tab' : 'text:tab-stop';

	return $element->appendChild($tabtag);
	}

#------------------------------------------------------------------------------

sub	createFrameElement
	{
	my $self	= shift;
	my %opt		= @_;
	my %attr	= ();

	$attr{'draw:name'} = $opt{'name'}; delete $opt{'name'};

	my $content_class = $self->contentClass;

	$attr{'draw:style-name'} = $opt{'style'}; delete $opt{'style'};
	if ($opt{'page'})
		{
		my $pg = $opt{'page'};
		if (ref $pg)
			{
			$opt{'attachment'} = $pg unless $opt{'attachment'};
			}
		elsif ($content_class eq 'text')
			{
			$opt{'attachment'} = $self->{'body'};
			$attr{'text:anchor-type'} = 'page';
			$attr{'text:anchor-page-number'} = $pg;
			}
		elsif 	(
				($content_class eq 'presentation')
					or
				($content_class eq 'drawing')
			)
			{
			my $n = $self->inputTextConversion($pg);
			$opt{'attachment'} = $self->getNodeByXPath
					("//draw:page[\@draw:name=\"$n\"]");
			}
		}
	delete $opt{'page'};

	my $tag = $opt{'tag'} || 'draw:frame'; delete $opt{'tag'};

	my $frame = OpenOffice::OODoc::XPath::new_element($tag);

	if ($opt{'position'})
		{
		$self->setObjectCoordinates($frame, $opt{'position'});
		delete $opt{'position'};
		}
	if ($opt{'size'})
		{
		$self->setObjectSize($frame, $opt{'size'});
		delete $opt{'size'};
		}
	if ($opt{'description'})
		{
		$self->setObjectDescription($frame, $opt{'description'});
		delete $opt{'description'};
		}
	if ($opt{'attachment'})
		{
		$frame->paste_first_child($opt{'attachment'});
		delete $opt{'attachment'};
		}

	foreach my $k (keys %opt)
		{
		$attr{$k} = $opt{$k} if ($k =~ /:/);
		}
	$self->setAttributes($frame, %attr);

	return $frame;
	}

sub	createFrame
	{
	my $self	= shift;
	return $self->createFrameElement(@_);
	}

#-----------------------------------------------------------------------------
# select an individual frame element by name

sub	selectFrameElementByName
	{
	my $self	= shift;
	my $text	= $self->inputTextConversion(shift);
	my $tag		= shift || 'draw:frame';
	return $self->selectNodeByXPath
			("//$tag\[\@draw:name=\"$text\"\]", @_);
	}

#-----------------------------------------------------------------------------
# gets frame element (name or ref, with type checking)

sub	getFrameElement
	{
	my $self	= shift;
	my $frame	= shift;
	return undef unless defined $frame;
	my $tag		= shift || 'draw:frame';

	my $element	= undef;
	if (ref $frame)
		{
		$element = $frame;
		}
	else
		{
		if ($frame =~ /^[\-0-9]*$/)
			{
			return $self->getElement("//$tag", $frame, @_);
			}
		else
			{
			return $self->selectFrameElementByName
				($frame, $tag, @_);
			}
		}
	}

#------------------------------------------------------------------------------

sub	getFrameList
	{
	my $self	= shift;
	return $self->getDescendants('draw:frame', shift);
	}

#------------------------------------------------------------------------------

sub	frameStyle
	{
	my $self	= shift;
	my $frame	= $self->getFrameElement(shift) or return undef;
	my $style	= shift;
	my $attr	= 'draw:style-name';
	return (defined $style) ?
		$self->setAttribute($frame, $attr => shift)	:
		$self->getAttribute($frame, $attr);
	}

#------------------------------------------------------------------------------
# replaces any previous content of an existing element by a given text
# without processing other than encoding

sub	setFlatText
	{
	my $self	= shift;
	my $path	= shift;
	my $element     = ref $path ?
	                        $path     :
	                        $self->OpenOffice::OODoc::XPath::getElement
	                                                ($path, shift);
	return undef unless $element;
	my $text	= shift;

	my $t		= $self->inputTextConversion($text);
	return undef unless defined $t;

	$element->set_text($t);
	return $text;
	}

#------------------------------------------------------------------------------
# replaces any previous content of an existing element by a given text
# processing tab stops and line breaks

sub	setText
	{
	my $self	= shift;
	my $path	= shift;
	my $element     = ref $path ?
	                        $path     :
	                        $self->OpenOffice::OODoc::XPath::getElement
	                                                ($path, shift);
	return undef unless $element;

	my $text	= shift;
	return undef	unless defined $text;

	unless ($text)
		{
		$element->set_text($text); return $text;
		}
	return $self->setFlatText($element, $text) if $element->isTextNode;

	my $tabtag = $self->{'opendocument'} ? 'text:tab' : 'text:tab-stop';
	$element->set_text("");
	my @lines	= split "\n", $text;
	while (@lines)
		{
		my $line	= shift @lines;
		my @columns	= split "\t", $line;
		while (@columns)
			{
			my $column	=
				$self->inputTextConversion(shift @columns);
			unless ($self->{'multiple_spaces'})
				{
				$element->appendTextChild($column);
				}
			else
				{
				$self->processSpaces($element, $column);
				}
			$element->appendChild($tabtag) if (@columns);
			}
		$element->appendChild('text:line-break') if (@lines);
		}
	$element->normalize;
	return $text;
	}

#------------------------------------------------------------------------------
# extends the text of an existing element

sub	extendText
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $text	= shift;

	return undef	unless defined $text;

	my $element 	= $self->getElement($path, $pos);
	return undef	unless $element;

	my $offset	= shift;

	if (ref $text)
		{
		if ($text->isElementNode)
			{
			unless (defined $offset)
				{
				$text->paste_last_child($element);
				}
			else
				{
				$text->paste_within($element, $offset);
				}
			}
		return $text;
		}

	my $tabtag = $self->{'opendocument'} ? 'text:tab' : 'text:tab-stop';
	my @lines	= split "\n", $text;
	my $ref_node	= undef;
	while (@lines)
		{
		my $line	= shift @lines;
		my @columns	= split "\t", $line;
		while (@columns)
			{
			my $column	=
				$self->inputTextConversion(shift @columns);
			unless ($ref_node)
				{
				$ref_node = $element->insertTextChild
						($column, $offset);
				$ref_node = $ref_node->insertNewNode
						($tabtag, 'after')
					if (@columns);
				}
			else
				{
				my $tn = $self->createTextNode($column);
				$ref_node = $ref_node->insertNewNode
						($tn, 'after');
				$ref_node = $ref_node->insertNewNode
						($tabtag, 'after')
					if (@columns);
				}
			}
		if (@lines)
			{
			if ($ref_node)
				{
				$ref_node->insertNewNode
						('text:line-break', 'after');
				}
			else
				{
			 	$element->insertNewNode
						(
						'text:line-break',
						'within',
						$offset
						);
				}
			}
		}

	$element->normalize;
	return $text;
	}

#------------------------------------------------------------------------------
# converts the content of an element to flat text

sub	flatten
	{
	my $self	= shift;
	my $element	= shift || $self->{'context'};
	return $element->flatten;
	}

#------------------------------------------------------------------------------
# creates a new encoded text node

sub	newTextNode
	{
	my $self	= shift;
	my $text	= $self->inputTextConversion(shift)
	                or return undef;
	return OpenOffice::OODoc::Element->new('#PCDATA' => $text);
	}

#------------------------------------------------------------------------------
# gets decoded text without other processing

sub	getFlatText
	{
	my $self	= shift;
	my $path	= shift;
	my $element     = ref $path ?
	                        $path     :
	                        $self->OpenOffice::OODoc::XPath::getElement
	                                                ($path, @_);
	return undef unless $element;

	return $self->outputTextConversion($element->text);
	}

#------------------------------------------------------------------------------
# gets text in element by path (sub-element texts are concatenated)

sub	getText
	{
	my $self	= shift;
        my $path        = shift;
        my $element     = ref $path ?
                                $path   :
                                $self->OpenOffice::OODoc::XPath::getElement
                                                        ($path, @_);
        return undef unless $element;
        return $self->getFlatText($element) if $element->isTextNode;
	return undef	unless $element->isElementNode;
	
	my $text	= '';

	my $name	= $element->getName;

	if	($name =~ /^text:tab(|-stop)$/)	{ return "\t"; }
	if	($name eq 'text:line-break')	{ return "\n"; }
	if	($name eq 'text:s')
		{
		my $spaces = "";
		my $count = $element->att('text:c') || 1;
		while ($count > 0) { $spaces .= ' '; $count--; }
		return $spaces;
		}
	foreach my $node ($element->getChildNodes)
		{
		if ($node->isElementNode)
			{
			$text .= $self->getText($node);
			}
		else
			{
			$text .= $self->outputTextConversion($node->text);
			}
		}
	return $text;
	}

#------------------------------------------------------------------------------

sub	xpathInContext
	{
	my $self	= shift;
	my $path	= shift	|| "/";
	my $context	= shift || $self->{'context'};
	if ($context ne $self->{'xpath'})
		{
		$path =~ s/^\//\.\//;
		}
	return ($path, $context);
	}

#------------------------------------------------------------------------------

sub	getDescendants
	{
	my $self	= shift;
	my $tag		= shift;
	my $context	= shift || $self->{'context'};
	return $context->descendants($tag, @_);
	}

#------------------------------------------------------------------------------

sub     getTextNodes
        {
        my $self        = shift;
	my $path	= shift;
	my $element	= ref $path ? $path : $self->getElement($path, shift)
	                        or return undef;
        my $filter      = $self->inputTextConversion(shift);
        return $element->getTextDescendants($filter);
        }

#------------------------------------------------------------------------------
# brute XPath nodelist selection; allows any XML::XPath expression

sub	selectNodesByXPath
	{
	my $self	= shift;
	my ($p1, $p2)	= @_;
	my $path	= undef;
	my $context	= undef;
	if (ref $p1)	{ $context = $p1; $path = $p2; }
	else		{ $path = $p1; $context = $p2; }
	($path, $context) = $self->xpathInContext($path, $context);
	unless (ref $context)
	        {
	        warn    "[" . __PACKAGE__ . "::selectNodesByXPath] "    .
	                "Bad context argument\n";
	        return undef;
	        }
	return $context->get_xpath($path);
	}

#------------------------------------------------------------------------------
# like selectNodesByXPath, without variable context (direct XML::Twig method)

sub     get_xpath
        {
        my $self        = shift;
        return $self->{'xpath'}->get_xpath(@_);
        }

#------------------------------------------------------------------------------
# brute XPath single node selection; allows any XML::XPath expression

sub	selectNodeByXPath
	{
	my $self	= shift;
	my $p1		= shift;
	my $p2		= shift;
	my $offset	= shift || 0;
	my $path	= undef;
	my $context	= undef;
	if (ref $p1)	{ $context = $p1; $path = $p2; }
	else		{ $path = $p1; $context = $p2; }
	($path, $context) = $self->xpathInContext($path, $context);
	unless (ref $context)
	        {
	        warn    "[" . __PACKAGE__ . "::selectNodeByXPath] "    .
	                "Bad context argument\n";
	        return undef;
	        }

	return $context->get_xpath($path, $offset);
	}

#------------------------------------------------------------------------------
# brute XPath value extraction; allows any XML::XPath expression

sub	getXPathValue
	{
	my $self	= shift;
	my ($p1, $p2)	= @_;
	my $path	= undef;
	my $context	= undef;
	if (ref $p1)	{ $context = $p1; $path = $p2; }
	else		{ $path = $p1; $context = $p2; }
	($path, $context) = $self->xpathInContext($path, $context);
	unless (ref $context)
	        {
	        warn    "[" . __PACKAGE__ . "::getXPathValue] "    .
	                "Bad context argument\n";
	        return undef;
	        }
	return $self->outputTextConversion($context->findvalue($path, @_));
	}

#------------------------------------------------------------------------------
# create or update an xpath

sub	makeXPath
	{
	my $self	= shift;
	my $path	= shift;
	my $root	= undef;
	if (ref $path)
		{
		$root	= $path;
		$path	= shift;
		}
	else
		{
		$root	= $self->getRoot;
		}
	$path =~ s/^[\/ ]*//; $path =~ s/[\/ ]*$//;
	my @list	= split '/', $path;
	my $posnode	= $root;
	while (@list)
		{
		my $item	= shift @list;
		while (($item =~ /\[.*/) && !($item =~ /\[.*\]/))
			{
			my $cont = shift @list or last;
			$item .= ('/' . $cont);
			}
		next unless $item;
		my $node	= undef;
		my $name	= undef;
		my $param	= undef;
		$item =~ s/\[(.*)\] *//;
		$param = $1;
		$name = $item; $name =~ s/^ *//; $name =~ s/ *$//;
		my %attributes = ();
		my $text = undef;
		my $indice = undef;
		if ($param)
			{
			my @attrlist = [];
			$indice = undef;
			$param =~ s/^ *//; $param =~ s/ *$//;
			$param =~ s/^@//;
			@attrlist = split /@/, $param;
			foreach my $a (@attrlist)
				{
				next unless $a;
				$a =~ s/^ *//;
				my $tmp = $a;
				$tmp =~ s/ *$//;
				if ($tmp =~ /^\d*$/)
					{
					$indice = $tmp;
					next;
					}
				if ($a =~ s/^\"(.*)\".*/$1/)
					{
					$text = $1; next;
					}
				if ($a =~ /^=/)
					{
					$a	=~ s/^=//;
					$a	=~ '^"(.*)"$';
					$text	= $1 ? $1 : $a;
					next;
					}
				$a =~ s/^@//;
				my ($attname, $attvalue) = split '=', $a;
				next unless $attname;
				if ($attvalue)
					{
					$attvalue =~ '"(.*)"';
					$attvalue = $1 if $1;
					}
				$attname =~ s/^ *//; $attname =~ s/ *$//;
				$attributes{$attname} = $attvalue;
				}
			}
		if (defined $indice)
			{
			$node = $self->getNodeByXPath
					($posnode, "$name\[$indice\]");
			}
		else
			{
			$node	=
				$self->getChildElementByName($posnode, $name);
			}
		if ($node)
			{
			$self->setAttributes($node, %attributes);
			$self->setText($node, $text)	if (defined $text);
			}
		else
			{
			$node = $self->appendElement
					(
					$posnode, $name,
					text		=> $text,
					attributes	=> {%attributes}
					);
			}
		if ($node)	{ $posnode = $node;	}
		else		{ return undef;		}
		}
	return $posnode;
	}

#------------------------------------------------------------------------------
# selects element by path and attribute

sub	selectElementByAttribute
	{
	my $self	= shift;
	my $path	= shift         or return undef;
	my $key		= shift         or return undef;
	my $arg3        = shift;
	
	my $xp  = undef;
	if (defined $arg3 && ! ref $arg3)       # arg3 = value
	        {
	        my $value = $self->inputTextConversion($arg3);
	        $xp = "//$path\[\@$key=\"$value\"\]";
	        }
	else                                    # arg3 = undef or context
	        {
	        $xp = "//$path\[\@$key\]" ; unshift @_, $arg3;
	        }
        
        return $self->selectNodeByXPath($xp, @_);
	}

#------------------------------------------------------------------------------

sub     selectElementByIdentifier
        {
        my $self        = shift;
        
        return $self->selectElementByAttribute(shift, $ELT_ID, @_);
        }

#------------------------------------------------------------------------------
# selects list of elements by path and attribute

sub	selectElementsByAttribute
	{
	my $self	= shift;
	my $path	= shift         or return undef;
	my $key		= shift         or return undef;
	my $arg3        = shift;
	
	my $xp  = undef;
	if (defined $arg3 && ! ref $arg3)       # arg3 = value
	        {
	        my $value = $self->inputTextConversion($arg3);
	        $xp = "//$path\[\@$key=\"$value\"\]";
	        }
	else                                    # arg3 = undef or context
	        {
	        $xp = "//$path\[\@$key\]" ; unshift @_, $arg3;
	        }
        

	return wantarray ?      $self->selectNodesByXPath($xp, @_)      :
	                        $self->selectNodeByXPath($xp, @_);
	}

#------------------------------------------------------------------------------
# get a list of elements matching a given path and an optional content pattern

sub	findElementList
	{
	my $self	= shift;
	my $path	= shift;
	my $pattern	= shift;
	my $replace	= shift;
	my $context	= shift;

	return undef unless $path;

	my @result	= ();

	($path, $context) = $self->xpathInContext($path, $context);
	foreach my $n ($context->findnodes($path))
		{
		push @result,
		    [ $self->findDescendants($n, $pattern, $replace, @_) ];
		}

	return @result;
	}

#------------------------------------------------------------------------------
# get a list of elements matching a given path and an optional content pattern
# without replacement operation, and from an optional context node

sub	selectElements
	{
	my $self	= shift;
	my $path	= shift;
	my $context	= $self->{'context'};
	if (ref $path)
		{
		$context	= $path;
		$path		= shift;
		}
	my $filter	= shift;

	my @candidates	= $self->selectNodesByXPath($context, $path);
	return @candidates	unless $filter;

	my @result	= ();
	while (@candidates)
		{
		my $node = shift @candidates;
		push @result, $node
			if $self->_search_content($node, $filter, @_, $node);
		}
	return @result;
	}

#------------------------------------------------------------------------------
# get the 1st element matching a given path and on optional content pattern

sub	selectElement
	{
	my $self	= shift;
	my $path	= shift;
	my $context	= $self->{'context'};
	if (ref $path)
		{
		$context	= $path;
		$path		= shift;
		}
	return undef	unless $path;
	my $filter	= shift;

	my @candidates	= $self->selectNodesByXPath($context, $path);
	return $candidates[0]	unless $filter;

	while (@candidates)
		{
		my $node = shift @candidates;
		return $node
			if $self->_search_content($node, $filter, @_, $node);
		}
	return undef;
	}

#------------------------------------------------------------------------------
# gets the descendants of a given node, with optional in fly search/replacement

sub	findDescendants
	{
	my $self	= shift;
	my $node	= shift;
	my $pattern	= shift;
	my $replace	= shift;

	my @result		= ();

	my $n	= $self->selectNodeByContent($node, $pattern, $replace, @_);
	push @result, $n	if $n;
	foreach my $m ($node->getChildNodes)
		{
		push @result,
		    [ $self->findDescendants($m, $pattern, $replace, @_) ];
		}

	return @result;
	}

#------------------------------------------------------------------------------
# search & replace text in an individual node

sub	selectNodeByContent
	{
	my $self	= shift;
	my $node	= shift;
	my $pattern	= shift;
	my $replace	= shift;

	return $node	unless $pattern;
	my $l	= $node->text;

	return undef	unless $l;

	unless (defined $replace)
		{
		return ($l =~ /$pattern/) ? $node : undef;
		}
	else
		{
		if (ref $replace)
			{
			unless
			  ($l =~ s/($pattern)/&$replace(@_, $node, $1)/eg)
				{
				return undef;
				}
			}
		else
			{
			unless ($l =~ s/$pattern/$replace/g)
				{
				return undef;
				}
			}
		$node->set_text($l);
		return $node;
		}
	}

#------------------------------------------------------------------------------
# gets the text content of a nodelist

sub	getTextList
	{
	my $self	= shift;
	my $path	= shift;
	my $pattern	= shift;
	my $context	= shift;

	return undef unless $path;

	($path, $context) = $self->xpathInContext($path, $context);
	my @nodelist = $context->findnodes($path);
	my @text = ();

	foreach my $n (@nodelist)
		{
		my $l	= $self->outputTextConversion($n->string_value);
		push @text, $l if ((! defined $pattern) || ($l =~ /$pattern/));
		}

	return wantarray ? @text : join "\n", @text;
	}

#------------------------------------------------------------------------------
# gets the attributes of an element in the key => value form

sub	getAttributes
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;

	my $node	= $self->getElement($path, $pos, @_);
	return undef	unless $path;

	my %attributes	= ();
	my $aa		= $node->atts(@_);
	my %atts	= %{$aa} if $aa;
	foreach my $a (keys %atts)
		{
		$attributes{$a}	= $self->outputTextConversion($atts{$a});
		}

	return %attributes;
	}

#------------------------------------------------------------------------------
# gets the value of an attribute by path + name

sub	getAttribute
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $name	= shift or return undef;

	my $node	= $self->getElement($path, $pos, @_);
	unless ($name =~ /:/)
	        {
	        my $prefix = $node->ns_prefix;
	        $name = $prefix . ':' . $name   if $prefix;
	        }
	$name =~ s/ /-/g;
	return	$self->outputTextConversion($node->att($name));
	}

#------------------------------------------------------------------------------
# set/replace a list of attributes in an element

sub	setAttributes
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my %attr	= @_;

	my $node	= $self->getElement($path, $pos, $attr{'context'});
	return undef	unless $node;
	my $prefix      = $node->ns_prefix();

	foreach my $k (keys %attr)
		{
		my $att_name = $k;
		$att_name =~ s/ /-/g;
		if (!($k =~ /:/) && $prefix)
		        {
		        $att_name = $prefix . ':' . $att_name;
		        }
		if (defined $attr{$k})
		    {
		    $node->set_att
		    		(
				$att_name,
				$self->inputTextConversion($attr{$k})
				);
		    }
		else
		    {
		    $node->del_att($att_name) if $node->att($att_name);
		    }
		}

	return %attr;
	}

#------------------------------------------------------------------------------
# set/replace a single attribute in an element

sub	setAttribute
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;

	my $attribute	= shift or return undef;
	my $value	= shift;
	my $node	= $self->getElement($path, $pos, @_)
		or return undef;

        $attribute =~ s/ /-/g;
        unless ($attribute =~ /:/)
                {
                my $prefix = $node->ns_prefix;
                $attribute = $prefix . ':' . $attribute if $prefix;                
                }
	if (defined $value)
		{
		$node->set_att
			(
			$attribute,
			$self->inputTextConversion($value)
			);
		}
	else
		{
		$node->del_att($attribute) if $node->att($attribute);
		}

	return $value;
	}

#------------------------------------------------------------------------------
# removes an attribute in element

sub	removeAttribute
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $name	= shift or return undef;

	my $node	= $self->getElement($path, $pos, @_)
	                        or return undef;

        unless ($name =~ /:/)
                {
                my $prefix = $node->ns_prefix;
                $name = $prefix . ':' . $name   if $prefix;
                }
	return $node->del_att($name) if $node->att($name);
	}

#------------------------------------------------------------------------------
# replicates an existing element, provided as an XPath ref or an XML string

sub	replicateElement
	{
	my $self	= shift;
	my $proto	= shift;
	my $position	= shift;
	my %options	= @_;

	unless ($proto && ref $proto && $proto->isElementNode)
		{
		warn "[" . __PACKAGE__ . "::replicateElement] No prototype\n";
		return undef;
		}

	$position	= 'end'	unless $position;

	my $element		= $proto->copy;
	$self->setAttributes($element, %{$options{'attribute'}});

	if	(ref $position)
		{
		if (! $options{'position'})
			{
			$element->paste_last_child($position);
			}
		elsif ($options{'position'} eq 'before')
			{
			$element->paste_before($position);
			}
		elsif ($options{'position'} eq 'after')
			{
			$element->paste_after($position);
			}
		elsif ($options{'position'} ne 'free')
			{
			warn	"[" . __PACKAGE__ . "::replicateElement] " .
				"No valid attachment option\n";
			}
		}
	elsif	($position eq 'end')
		{
		$element->paste_last_child($self->{'xpath'}->root);
		}
	elsif	($position eq 'body')
		{
		$element->paste_last_child($self->getBody);
		}

	return $element;
	}

#------------------------------------------------------------------------------
# create an element, just with a mandatory name and an optional text
# the name can have the namespace:name form
# if the $name argument is a '<.*>' string, it's processed as XML and
# the new element is completely generated from it

sub	createElement
	{
	my $self	= shift;
	my $name	= shift;
	my $text	= shift;

	my $element = OpenOffice::OODoc::XPath::new_element($name, @_);
	unless ($element)
		{
		warn	"[" . __PACKAGE__ . "::createElement] "	.
			"Element creation failure\n";
		return undef;
		}

	$self->setText($element, $text)		if defined $text;

	return $element;
	}

#------------------------------------------------------------------------------
# replaces an element by another one
# the new element is inserted before the old one,
# then the old element is removed.
# the new element can be inserted by copy (default) or by reference
# return = new element if success, undef if failure

sub	replaceElement
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $new_element	= shift;
	my %options	=
			(
			mode		=> 'copy',
			@_
			);
	unless ($new_element)
		{
		warn	"[" . __PACKAGE__ . "::replaceElement] " .
			"Missing new element\n";
		return undef;
		}
	unless (ref $new_element)
		{
		$new_element = $self->createElement($new_element);
		$options{'mode'} = 'reference';
		}
	unless ($new_element && $new_element->isElementNode)
		{
		warn	"[" . __PACKAGE__ . "::replaceElement] " .
			"No valid replacement\n";
		return undef;
		}

	my $result	= undef;

	my $old_element	= $self->getElement
			($path, $pos, $options{'context'});
	unless ($old_element)
		{
		warn	"[" . __PACKAGE__ . "::replaceElement] " .
			"Non existing element to be replaced\n";
		return undef;
		}
	if	(! $options{'mode'} || $options{'mode'} eq 'copy')
		{
		$result = $new_element->copy;
		$result->replace($old_element);
		return $result;
		}
	elsif	($options{'mode'} && $options{'mode'} eq 'reference')
		{
		$result = $self->insertElement
					(
					$old_element,
					$new_element,
					position	=> 'before'
					);
		$old_element->delete;
		return $result;
		}
	else
		{
		warn	"[" . __PACKAGE__ . "::replaceElement] " .
			"Unknown option\n";
		}
	return undef;
	}

#------------------------------------------------------------------------------
# appends a new or existing child element to any existing element

sub	appendElement
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $name	= shift;
	my %opt		= @_;
	$opt{'attribute'} = $opt{'attributes'} unless ($opt{'attribute'});

	return undef	unless $name;
	my $element	= undef;

	unless (ref $name)
		{
		$element	= $self->createElement($name, $opt{'text'});
		}
	else
		{
		$element	= $name;
		$self->setText($element, $opt{'text'})	if $opt{'text'};
		}
	return undef	unless $element;
	my $parent	= $self->getElement
			($path, $pos, $opt{'context'});
	unless ($parent)
		{
		warn	"[" . __PACKAGE__ .
			"::appendElement] Position not found\n";
		return undef;
		}
	$element->paste_last_child($parent);
	$self->setAttributes($element, %{$opt{'attribute'}});

	return $element;
	}

#-----------------------------------------------------------------------------
# append an element to the document body

sub	appendBodyElement
	{
	my $self	= shift;

	return $self->appendElement($self->{'body'}, @_);
	}

#------------------------------------------------------------------------------
# appends a list of children to an existing element

sub	appendElements
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $parent	= $self->getElement($path, $pos) or return undef;
	my @children	= @_;
	foreach my $child (@children)
		{
		$parent->appendChild($child);
		}
	return $parent;
	}

#------------------------------------------------------------------------------
# cuts a set of existing elements and pastes them as children of a given one

sub	moveElements
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $parent	= $self->getElement($path, $pos) or return undef;
	$parent->pickUpChildren(@_);
	return $parent;
	}

#------------------------------------------------------------------------------
# selects a text node in a given element according to offset & expression

sub     textIndex
        {
        my $self        = shift;
        my $path        = shift;
        my $element     = (ref $path) ? $path : $self->getElement($path, shift)
                        or return undef;
	my %opt         = @_;

        my $offset      = $opt{'offset'};
        my $way         = $opt{'way'} || 'forward';
        if (defined $offset && $offset < 0)
                {
                $way = 'backward';
                }
        $offset = -abs($offset) if defined $offset && $way eq 'backward';

        my $start_mark  = $opt{'start_mark'};
        my $end_mark    = $opt{'end_mark'};

        my $expr        = undef;
        if (defined $opt{'after'})
                {
                $expr = $opt{'after'};
                delete @opt{qw(before replace capture content)};
                }
        elsif (defined $opt{'before'})
                {
                $expr = $opt{'before'};
                delete @opt{qw(replace capture content)};
                }
        else
                {
                $expr = $opt{'content'} || $opt{'replace'} || $opt{'capture'};
                }
        $expr           = $self->inputTextConversion($expr);

        my $node        = undef;
        my $node_text   = undef;
        my $node_length = undef;
        my $found       = undef;
        my $end_pos     = undef;
        my $match       = undef;

        if ($way ne 'backward')         # positive offset, forward
                {
                if ($element->isTextNode)
                        {
                        $node = $element;
                        }
                elsif ($start_mark)
                        {
                        unless($start_mark->isTextNode)
                                {
                                my $n   = $start_mark->last_descendant;
                                $start_mark = $n        if $n;
                                $node   = $n->next_elt($element, '#PCDATA');
                                }
                        else
                                {
                                $node   = $start_mark;
                                }
                        }
                else
                        {
                        $node = $element->first_descendant('#PCDATA');
                        }
                if ($end_mark && ! $node->before($end_mark))
                        {
                        $node = undef;
                        }
                ($node_length, $node_text) = $node->textLength  if $node;
                FORWARD_LOOP: while ($node && !defined $found)
                        {
                        if ($end_mark && ! $node->before($end_mark))
                                {
                                $node = undef;
                                last;
                                }
                        if (defined $offset && ($offset > $node_length))
                                {                       # skip node
                                $offset -= $node_length;
                                $node = $node->next_elt($element, '#PCDATA');
                                ($node_length, $node_text) = $node->textLength
                                        if $node;
                                }

                        elsif (defined $expr)
                                {                       # look for substring
                                my $text = $node->text() || "";
                                if (defined $offset && $offset > 0)
                                        {
                                        $text = substr($text, $offset);
                                        }
                                if ($text =~ /($expr)/)
                                        {
                                        $found = length($`);
                                        $found += $offset if defined $offset;
                                        $end_pos = $found + length($&);
                                        $match = $1;
                                        }
                                unless (defined $found)
                                        {
                                        $offset = undef;
                                        $node   = $node->next_elt
                                                        ($element, '#PCDATA');
                                        }
                                }
                        else                              # selected by offset
                            {
                            $found = $offset || 0;
                            }
                        }
                }
        else                            # negative offset, backward
                {
                if ($element->isTextNode)
                        {
                        $node = $element;
                        }
                elsif ($start_mark)
                        {
                        unless ($start_mark->isTextNode)
                                {
                                $node   = $start_mark->prev_elt('#PCDATA');
                                }
                        else
                                {
                                $node   = $start_mark;
                                }
                        }
                else
                        {
                        $node   = $element->last_descendant('#PCDATA');
                        }
                if ($end_mark)
                        {
                        my $n = $end_mark->last_descendant;
                        $end_mark = $n        if $n;
                        $node = undef if
                                ($end_mark && ! $node->after($end_mark));
                        }
                ($node_length, $node_text) = $node->textLength  if $node;
                BACKWARD_LOOP: while ($node && !defined $found)
                        {
                        if ($end_mark && ! $node->after($end_mark))
                                {
                                $node = undef;
                                last;
                                }
                        ($node_length, $node_text) = $node->textLength;
                        if (defined $offset && (abs($offset) > $node_length))
                                {                       # skip node
                                $offset += $node_length;
                                $node = $node->prev_elt($element, '#PCDATA');
                                }
                        elsif (defined $expr)
                                {
                                my $text = $node->text() || "";
                                if (defined $offset && $offset < 0)
                                        {
                                        $text = substr($text, 0, $offset);
                                        }
                                my @r = ($text =~ m/($expr)/g);
                                if (@r)
                                        {
                                        $found = length($`);
                                        $end_pos = $found + length($&);
                                        $match = $1;
                                        }
                                unless (defined $found)
                                        {
                                        $offset = undef;
                                        $node   = $node->prev_elt
                                                        ($element, '#PCDATA');
                                        }
                                }
                        else                              # selected by offset
                                {
                                $found = $offset || 0;
                                }
                        }
                }

        return ($node, $found, $end_pos, $match);
        }
 
#------------------------------------------------------------------------------
# creates new child elements in a given element and splits the content
# according to a regexp

sub     splitContent
        {
        my $self        = shift;
        my $path        = shift;
        my $pos         = (ref $path) ? undef : shift;
        my $context     = $self->getElement($path, $pos) or return undef;
        my $tag         = shift         or return undef;
        my $expr        = $self->inputTextConversion(shift);
        return undef unless defined $expr;        
        my %opt         = @_;

        my $prefix      = undef;        
        if ($tag =~ /(.*):/)
                {
                $prefix = $1 || 'text';
                }
        else
                {
                $prefix = $context->ns_prefix() || 'text';
                $tag = $prefix . ':' . $tag;
                }

        my %attr        = ();
        foreach my $k (keys %opt)
                {
                my $a = $self->inputTextConversion($opt{$k});
                $k = $prefix . ':' . $k         unless $k =~ /:/;
                $attr{$k} = $a;
                }
        %opt            = ();
        
        return $context->mark("($expr)", $tag, { %attr });
        }

#------------------------------------------------------------------------------
# creates a child element in place within an existing element
# at a given position or before/after a given substring

sub     setChildElement 
        {
        my $self        = shift;
        my $path        = shift;
        my $node        = (ref $path) ? $path : $self->getElement($path, shift)
                        or return undef;
        my $name        = shift or return undef;
        my %opt         = @_;
        if (defined $opt{'text'})
                {
                $opt{'replace'} = $opt{'capture'}
                                unless defined $opt{'replace'};
                delete $opt{'capture'};
                }
        my $newnode     = undef;
        my $function    = undef;
    
        if (ref $name)
                {
                if      ((ref $name) eq 'CODE')
                        {
                        $function   = $name;
                        $name       = undef;
                        }
                else
                        {
                        $newnode    = $name;
                        }
                }
        else
                {
                unless ($name =~ /:/ || $name =~ /^#/)
                        {
                        my $prefix = $node->ns_prefix() || 'text';
                        $name = $prefix . ':' . $name;
                        }
                $newnode = OpenOffice::OODoc::XPath::new_element($name);
                }

       	my $offset = $opt{'offset'} || 0;
	if (lc($offset) eq 'end')
		{
		unless ($function)
		        {
		        $newnode->paste_last_child($node);
		        }
		else
		        {
		        $newnode = &$function($self, $node, 'end');
		        }
		}
	elsif (lc($offset) eq 'start')
	        {
	        unless ($function)
	                {
	                $newnode->paste_first_child($node);
	                }
                else
                        {
                        $newnode = &$function($self, $node, 'start');
                        }
	        }
	else
		{
                my ($text_node, $start_pos, $end_pos, $match) =
                                        $self->textIndex($node, %opt);
                if ($text_node)
                        {
                        if (defined $opt{'replace'} || defined $opt{'capture'})
                                {
                                my $t = $text_node->text;
                                substr  (
                                        $t, $start_pos, $end_pos - $start_pos,
                                        ""
                                        );
                                $text_node->set_text($t);
                                unless ($function)
                                        {
                                        $newnode->paste_within
                                                ($text_node, $start_pos);
                                        $newnode->set_text($match)
                                                if defined $opt{'capture'};
                                        }
                                else
                                        {
                                        $newnode = &$function
                                                        (
                                                        $self,
                                                        $text_node,
                                                        $start_pos,
                                                        $match
                                                        );
                                        }
                                }
                        else
                                {
                                my $p = defined $opt{'after'} ?
                                                $end_pos : $start_pos;
                                unless ($function)
                                        {
                                        $newnode->paste_within($text_node, $p);
                                        }
                                else
                                        {
                                        $newnode = &$function
                                                        (
                                                        $self,
                                                        $text_node,
                                                        $p,
                                                        $match
                                                        );
                                        }
                                }
                        }
                else
                        {
                        return undef;
                        }
		}

        if ($newnode)
                {
                $self->setAttributes($newnode, %{$opt{'attributes'}});
                $self->setText($newnode, $opt{'text'})
                                unless is_true($opt{'no_text'});
                }
        return $newnode;
        }

#------------------------------------------------------------------------------
# create successive child elements

sub     setChildElements
        {
        my $self        = shift;
        my $path        = shift;
        my $pos         = (ref $path) ? undef : shift;
        my $element     = $self->getElement($path, $pos) or return undef;
        my $name        = shift or return undef;
        my %opt         = @_;

        my @elements    = ();
        my $node        = $self->setChildElement($element, $name, %opt);
        push @elements, $node if $node;

        if (defined $opt{'text'})
                {
                $opt{'replace'} = $opt{'capture'}
                                unless defined $opt{'replace'};
                delete $opt{'capture'};
                }

        delete $opt{'attributes'};
        delete $opt{'text'};
        delete $opt{'offset'} if
                (
                defined $opt{'after'}   ||
                defined $opt{'before'}  ||
                defined $opt{'replace'} ||
                defined $opt{'capture'}
                ); 
        $opt{'offset'} = 1 if
                (
                    ($opt{'way'} ne 'backward' && defined $opt{'before'})
                        ||
                    ($opt{'way'} eq 'backward' && defined $opt{'after'})
                );

        while ($node)
                {
                my $arg = ref($name) eq 'CODE' ? $name : $node->copy;
                $node = $self->setChildElement
                                ($element, $arg, %opt, start_mark => $node);
                push @elements, $node if $node;
                }

        return @elements;
        }

#------------------------------------------------------------------------------

sub     markElement
        {
        my $self        = shift;
        my $context     = shift         or return undef;
        my $tag         = shift;
        my $expression  = $self->inputTextConversion(shift);
        my %attr        = @_;
        
        return $context->mark("($expression)", $tag, { %attr });
        }

#------------------------------------------------------------------------------
# inserts a new element before or after a given node

sub	insertElement
	{
	my $self	= shift;
	my $path	= shift;
	my $pos		= (ref $path) ? undef : shift;
	my $name	= shift;
	my %opt		= @_;
	$opt{'attributes'} = $opt{'attribute'} unless $opt{'attributes'};

	return undef	unless $name;
	my $element	= undef;
	unless (ref $name)
		{
		$element	= $self->createElement($name, $opt{'text'});
		}
	else
		{
		$element	= $name;
		$self->setText($element, $opt{'text'})	if $opt{'text'};
		}
	return undef	unless $element;

	my $posnode	= $self->getElement($path, $pos, $opt{'context'});
	unless ($posnode)
		{
		warn "[" . __PACKAGE__ . "::insertElement] Unknown position\n";
		return undef;
		}

	if ($opt{'position'})
		{
		if ($opt{'position'} eq 'after')
			{
			$element->paste_after($posnode);
			}
		elsif ($opt{'position'} eq 'before')
			{
			$element->paste_before($posnode);
			}
		elsif ($opt{'position'} eq 'within')
			{
			my $offset = $opt{'offset'} || 0;
			$element->paste_within($posnode, $offset);
			}
		else
			{
			warn	"[" . __PACKAGE__ . "::insertElement] "	.
				"Invalid $opt{'position'} option\n";
			return undef;
			}
		}
	else
		{
		$element->paste_before($posnode);
		}

	$self->setAttributes($element, %{$opt{'attributes'}});

	return $element;
	}

#------------------------------------------------------------------------------
# removes the given element & children

sub	removeElement
	{
	my $self	= shift;

	my $e	= $self->getElement(@_);
	return undef	unless $e;
	return $e->delete;
	}

#------------------------------------------------------------------------------
# cuts the given element & children (to be pasted elsewhere)

sub	cutElement
	{
	my $self	= shift;

	my $e	= $self->getElement(@_);
	return undef	unless $e;
	$e->cut;

	return $e;
	}

#-----------------------------------------------------------------------------
# splits a text element at a given offset

sub	splitElement
	{
	my $self	= shift;
	my $path	= shift;
	my $old_element	=
		(ref $path) ? $path : $self->getElement($path, shift);
	my $offset	= shift;

	my $new_element = $old_element->split_at($offset);
	$new_element->set_atts($old_element->atts);
	return wantarray ? ($old_element, $new_element) : $new_element;
	}

#------------------------------------------------------------------------------
# get/set ODF element identifier

sub     getIdentifier
        {
        my $self        = shift;
	my $path	= shift;
	my $element	=
		(ref $path) ? $path : $self->getElement($path, shift);
	return $self->outputTextConversion($element->getID());
        }

sub     setIdentifier
        {
        my $self        = shift;
	my $path	= shift;
	my $element	=
		(ref $path) ? $path : $self->getElement($path, shift);
	my $value       = shift;
	return (defined $value) ?
	        $self->inputTextConversion($element->setID($value))     :
	        $self->removeIdentifier($element);
        }

sub     identifier
        {
        my $self        = shift;
	my $path	= shift;
	my $element	=
		(ref $path) ? $path : $self->getElement($path, shift);
	my $value       = shift;
        return (defined $value) ?
                $self->setIdentifier($element, $value)  :
                $self->getIdentifier($element);
        }

sub     removeIdentifier
        {
        my $self        = shift;
	my $path	= shift;
	my $element	=
		(ref $path) ? $path : $self->getElement($path, shift);
	return $element->setID();
        }

sub     getElementName
        {
        my $self        = shift;
        my $path        = shift;
	my $element	=
		(ref $path) ? $path : $self->getElement($path, shift);
        my $attr        = $element->ns_prefix() . ':name';
        return $self->getAttribute($element, $attr);
        }

sub     setElementName
        {
        my $self        = shift;
        my $path        = shift;
	my $element	=
		(ref $path) ? $path : $self->getElement($path, shift);
        my $attr        = $element->ns_prefix() . ':name';
        return $self->setAttribute($element, $attr => shift);       
        }

sub     elementName
        {
        my $self        = shift;
	my $path	= shift;
	my $element	=
		(ref $path) ? $path : $self->getElement($path, shift);
	my $value       = shift;
        return (defined $value) ?
                $self->setElementName($element, $value)  :
                $self->getElementName($element);       
        }

#------------------------------------------------------------------------------
# some extensions for XML Twig elements
package	OpenOffice::OODoc::Element;
our @ISA	= qw ( XML::Twig::Elt );
#------------------------------------------------------------------------------

BEGIN   {
	*identifier             = *ID;
	*getPrefix              = *XML::Twig::Elt::ns_prefix;
	*getNodeValue           = *XML::Twig::Elt::text;
	*getValue               = *XML::Twig::Elt::text;
	*setNodeValue           = *XML::Twig::Elt::set_text;
	*getAttribute           = *XML::Twig::Elt::att;
	*setName                = *XML::Twig::Elt::set_tag;
	*getParentNode          = *XML::Twig::Elt::parent;
	*getDescendantTextNodes = *getTextDescendants;
	*dispose                = *XML::Twig::Elt::delete;
        }

sub	hasTag
	{
	my $node	= shift;
	my $name	= $node->getName;
	my $value	= shift;
	return ($name && ($name eq $value)) ? 1 : undef;
	}
	
sub	isFrame
	{
	my $node	= shift;
	return $node->hasTag('draw:frame');
	}

sub	getLocalPosition
	{
	my $node	= shift;
	my $tag		= (shift || $node->getName) or return undef;
	my $xpos	= $node->pos($tag);
	return defined $xpos ? $xpos - 1 : undef;
	}

sub     selectChildElements
        {
        my $node        = shift;
        my $filter      = shift;
        my $condition   = ref $filter ? $filter : qr($filter);
        return $node->children($condition);
        }

sub	selectChildElement
	{
	my $node	= shift;
	my $filter	= shift;
	my $pos		= shift || 0;

	my $count	= 0;
	my $fc = $node->first_child;
	return $fc unless defined $filter;
	my $name = $fc->name if $fc;
	while ($fc)
		{
		if ($name && ($name =~ /$filter/))
			{
			return $fc if ($count >= $pos);
			$count++;
			}
		$fc = $fc->next_sibling;
		$name = $fc->name if $fc;
		}
	return undef;
	}

sub	getFirstChild
	{
	my $node	= shift;
	my $fc = $node->first_child(@_);
	my $name = $fc->name if $fc;
	while ($name && ($name =~ /^#/))
		{
		$fc = $fc->next_sibling(@_);
		$name = $fc->name if $fc;
		}
	return $fc;
	}

sub	getLastChild
	{
	my $node	= shift;
	my $lc = $node->last_child(@_);
	my $name = $lc->name;
	while ($name && ($name =~ /^#/))
		{
		$lc = $lc->prev_sibling(@_);
		$name = $lc->name;
		}
	return $lc;
	}

sub     getChildrenTextNodes
        {
        my $node        = shift;
        return $node->children('#PCDATA');
        }

sub     getChildTextNode
        {
        my $node        = shift;
        my $pos         = shift || 0;
        my @children    = $node->children('#PCDATA');
        return $children[$pos];
        }

sub	getTextDescendants
	{
	my ($node, $filter)     = @_;
	return  defined $filter ?
	        $node->get_xpath('#PCDATA[string()=~/' . $filter . '/]') :
	        $node->descendants('#PCDATA');
	}

sub     textLength      # length of a text node
        {
        my $node        = shift;
        my $text        = $node->text;
        my $length      = length($text);
        return wantarray ? ($length, $text) : $length;
        }

sub	appendChild
	{
	my $node	= shift;
	my $child	= shift;
	unless (ref $child)
		{
		$child = OpenOffice::OODoc::XPath::new_element($child, @_);
		}
	return $child->paste_last_child($node);
	}

sub	pickUpChildren
	{
	my $parent	= shift;
	my @children	= @_;
	foreach my $child (@children)
		{
		$child->move(last_child => $parent);
		}
	return $parent;
	}

sub	insertNewNode
	{
	my $node	= shift;
	my $newnode	= shift or return undef;
	my $position	= shift;	# 'before', 'after', 'within', ...
	my $offset	= shift;
	unless (ref $newnode)
		{
		$newnode = OpenOffice::OODoc::XPath::new_element($newnode, @_);
		}
	if (defined $offset)
		{
		return $newnode->paste($position => $node, $offset);
		}
	else
		{
		return $newnode->paste($position => $node);
		}
	}

sub     insertNodes
        {
        my $node        = shift;
        my $offset      = shift;
        my $child       = shift         or return undef;
        $child->paste_within($node, $offset);
        my $count = 1;
        while (@_)
                {
                my $next_child = shift;
                $next_child->paste_after($child);
                $child = $next_child;
                $count++;
                }
        return $count;
        }

sub	replicateNode
	{
	my $node	= shift;
	my $number	= shift;
	$number = 1 unless defined $number;
	my $position	= shift || 'after';
	my $last_node	= $node;
	while ($number > 0)
		{
		my $newnode	= $node->copy;
		$newnode->paste($position => $last_node);
		$last_node	= $newnode;
		$number--;
		}
	return $last_node;
	}

sub	flatten
	{
	my $node	= shift;
	return $node->set_text($node->text);
	}

sub	appendTextChild
	{
	my $node	= shift;
	my $text	= shift;
	return undef unless defined $text;
	my $text_node	= OpenOffice::OODoc::Element->new('#PCDATA' => $text);
	return $text_node->paste_last_child($node);
	}

sub	insertTextChild
	{
	my $node	= shift;
	my $text	= shift;
	return undef unless defined $text;
	my $offset	= shift;
	return $node->appendTextChild($text) unless defined $offset;
	my $text_node	= OpenOffice::OODoc::Element->new('#PCDATA' => $text);
	return $offset > 0 ?
	        $text_node->paste_within($node, $offset)        :
	        $text_node->paste_first_child($node);
	}

sub	getAttributes
	{
	my $node	= shift;
	return %{$node->atts(@_) || {}};
	}

sub	setAttribute
	{
	my $node	= shift or return undef;
	my $attribute	= shift;
	my $value	= shift;
	if (defined $value)
		{
		return $node->set_att($attribute, $value, @_);
		}
	else
		{
		return $node->removeAttribute($attribute);
		}
	}

sub     setID
        {
        my $node        = shift;
        return $node->setAttribute($ELT_ID, shift);
        }

sub     getID
        {
        my $node        = shift;
        return $node->getAttribute($ELT_ID);
        }

sub     ID
        {
        my $node        = shift;
        my $new_id      = shift;
        return (defined $new_id) ? $node->setID($new_id) : $node->getID();
        }

sub	removeAttribute
	{
	my $node	= shift or return undef;
	my $attribute	= shift or return undef;
	return $node->att($attribute) ? $node->del_att($attribute) : undef;
	}

#------------------------------------------------------------------------------
1;
