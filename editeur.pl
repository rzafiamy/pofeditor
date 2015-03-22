#!/bin/perl

#________________________________________________
# Auteur: Rija ZAFIAMY
# Date : fevrier 2015
# Generateur de document openoffice
#________________________________________________

use OpenOffice::OODoc;

#________________________________________________
# Variables globales
#________________________________________________
my $countTab = 0;
my $countImg = 0;
my @contenu;

# creation d'un fichier
my $fichier = ooDocument
(
	file 	=> 'rapport.odt',
	create 	=> 'text',
	member	=> 'content'
);

$fichier->save;


# Ouverture du fichier
my $archive 	= ooFile('rapport.odt');

# instance objet Meta
my $meta 		= ooMeta
(
	archive => $archive 
);
# instance objet Document pour content
my $texte 		= ooDocument
(
	archive => $archive,
	member 	=> 'content'
);
# instance objet Document pour les styles
my $deco = ooDocument
(
	archive => $archive,
	member 	=> 'styles'
);
# definit le titre du document
$meta->title('Rapport');

# --------------------------------------
# Creation de style de page
#---------------------------------------
my $layout = $deco->pageLayout("Standard");
my $master = $deco->createMasterPage("StylePages",layout=>$layout);
                	
creer_style($ARGV[1]);

automate($ARGV[0]);

$archive->save;

#foreach my $e (@contenu)
#{
	
	#print $e."\n";
	#$l = <STDIN>;
#}
#_________________________________________
#
# Insere un titre de niveau 1-3 dans le document
#_________________________________________
sub creer_titre
{
	my ($niveau,$titre) = @_;
	$titre =~ s/_TAB_/\t/g;
	$texte->appendParagraph
	(
	 	style => 'StyleTitre'.$niveau,
		text  => $titre
	);
}

#_________________________________________
#
# Insere une tableau dans le document
#_________________________________________
sub creer_tableau
{
	
	my ($l,$c,@tab) = @_;
		
	my $feuille = $texte->appendTable('Table'.$countTab, $l, $c);

	my $i = 0;
	
	foreach my $li (@tab)
	{
		my $j = 0;
		
		foreach my $col (@$li)
		{
			$col =~ s/_TAB_//g;
			$texte->cellValue($feuille, $i,$j,$col);
			push(@contenu,$col);
			$j++;
		}
		$i++;
	}
	
	$texte->appendParagraph
	(
	 	style 	=> 'StyleParagraphe',
		text  	=> ""
	);
	
	$countTab++;
}

#_________________________________________
#
# Insere une liste dans le document
#_________________________________________
sub creer_liste
{
	
	my (@liste) = @_;
	
	foreach my $e (@liste)
	{
		$e =~ s/_TAB_//g;
		$e =~ s/_DIESE_/#/g;
		$e = ajouter_tabulation($e);
		my @items = split("_NL_",$e);
		
		foreach my $item (@items)
		{
			$texte->appendParagraph
			(
				style => 'StyleListe',
				text  => $item
			);
			push(@contenu,$item);
		}
	}
}

#_________________________________________
#
# Insere une image dans le document
#_________________________________________
sub creer_image
{
	
	my ($img) = @_;
	
	my @param = split(",",$img);
	my $p = creer_paragraphe("");
		
	$texte->createImageElement
                (
					"Image".$countImg,
					title    => $param[2],
					description     => $param[2],
					attachment      => $p,
					import          => $param[3],
					size 			=> $param[0]." cm,".$param[1]." cm",
					style			=> "StyleImage"
                );
	$countImg++;
}

#_________________________________________
#
# Insere un paragraphe dans le document
#_________________________________________
sub creer_paragraphe
{
	my ($paragraphe) = @_;
	
	$paragraphe =~ s/_TAB_//g;
	$paragraphe =~ s/_DIESE_/#/g;
	$paragraphe = ajouter_tabulation($paragraphe);
	my @parag = split("_NL_",$paragraphe);
	

	foreach my $p (@parag)
	{
		$texte->appendParagraph
		(
			style 	=> 'StyleParagraphe',
			text  	=> $p
		);
		push(@contenu,$p);
	}
	
	my $ret;
		
	$ret = $texte->appendParagraph
	(
	 	style 	=> 'StyleParagraphe',
		text  	=> ""
	);
	
	return $ret;
	
}

#_________________________________________
#
# Insere un paragraphe dans le document
#_________________________________________
sub creer_code
{
	my ($code) = @_;
	
	$code =~ s/_TAB_/\t/g;
	$code =~ s/_DIESE_/#/g;
	my @lignes = split("_NL_",$code);
	

	foreach my $c (@lignes)
	{
		$texte->appendParagraph
		(
			style 	=> 'StyleCode',
			text  	=> $c
		);
		push(@contenu,$c);
	}
	
	my $ret;
		
	$ret = $texte->appendParagraph
	(
	 	style 	=> 'StyleParagraphe',
		text  	=> ""
	);
	
	return $ret;
	
}

#_________________________________________
#
# Creer une en-tete de page
#_________________________________________
sub creer_EntetePage
{
	
	my ($titre) = @_;
	$titre = ajouter_tabulation($titre);
	my $p = $deco->createParagraph($titre,"StyleEntete");
	
	$deco->masterPageExtension($master,"header",$p);
	
	push(@contenu,$titre);
}

#_________________________________________
#
# Creer un pied de page
#_________________________________________
sub creer_PiedPage
{
	my ($titre) = @_;
	
	$titre = ajouter_tabulation($titre);
	my $p = $deco->createParagraph($titre,"StylePied");
	$deco->appendElement($p, $deco->textField("page-number"));
	$deco->masterPageExtension($master,"footer",$p);
	push(@contenu,$titre);
}

#_________________________________________
#
# Creer un en-tete
#_________________________________________
sub creer_saut_page
{
	
	my ($param) = @_;
		
	my $p = creer_paragraphe("");
	$texte->setPageBreak($p,style=> "xyz",page=>"StylePages");
}

#_________________________________________
#
# Rajoute des tabulations
#_________________________________________
sub ajouter_tabulation
{
	my ($in) = @_;
	
	my $ret = "";
	my $paragraphe = $in;
	
	while($paragraphe ne ""){
	
		if($paragraphe =~ /_TAB\(([1-9]+)\)/)
		{
			my $i = 0;
			my $tab = "";
			for(;$i < $1;$i++)
			{
				$tab = $tab."\t";
			}
			$ret = $ret.$`.$tab;
			$paragraphe = $';
		}
		else
		{
			if($ret ne "")
			{
				$ret = $ret.$paragraphe;
			}
			$paragraphe = "";
			
		}
	}
	if($ret eq "")
	{
		return $in;
	}	
	return $ret;
}
#_________________________________________
#
# Automate qui gère les entrees
# Est-ce :
# 	- un titre ?
#	- un paragraphe ?
#	- une image ?
#	- une liste ?
#	- un tableau ?
#_________________________________________
sub automate
{
	
	my ($filename) = @_;
	
	if($filename eq "-f")
	{
		open(FICHIER ,"<".$ARGV[1]) or die("Ereur ouverture fichier");
		
		$raw = "";
		
		while($ligne = <FICHIER>)
		{
			$ligne =~ s/\n/_NL_\n/g;
			$ligne =~ s/\[/\\\[/g;
			$ligne =~ s/\]/\\\]/g;
			$ligne =~ s/\{/\\\{/g;
			$ligne =~ s/\}/\\\}/g;
			print $ligne;
		}
		close(FICHIER);
	}
	else
	{
		open(FICHIER ,"<".$filename) or die("Ereur ouverture fichier modele");
		
		$raw = "";
		
		while($ligne = <FICHIER>)
		{
			chomp($ligne);
			$raw = $raw.$ligne;
		}
		
		close(FICHIER);	
		
		#pre-traitement
		$raw =~ s/#/_DIESE_/g;
		$raw =~ s/\t/_TAB_/g;
		$raw =~ s/(_TAB_)+\[/\[/g;# _TAB_[
		$raw =~ s/\](_TAB_)+/\]/g;# ]_TAB_
		
		$raw =~ s/(_TAB_)+\{/\{/g; # _TAB_{
		$raw =~ s/\}(_TAB_)+/\}/g; # }_TAB_
		
		$raw =~ s/([\s])[\s]{1,}/$1/g;
		
		$raw =~ s/([^\\]{1})([\{,])\{/$1$2:LLIGNE:/g; # forme {{ ou ,{
		$raw =~ s/([^\\]{1})\}([\},])/$1:RLIGNE:$2/g; # forme }, ou }}  
		
		$raw =~ s/([^\\]{1})\[/$1#LTAG#/g; #forme [
		$raw =~ s/([^\\]{1})\]:/$1#RTAG#:/g; #forme ]:
		
		$raw =~ s/:\{/#LBLOC#/g;
		$raw =~ s/([^\\]{1})\}/$1#RBLOC#/g;
		
		$raw =~ s/\\(.)/$1/g; # caracteres echappes
		$raw =~ s/_DEBUT_//g; # caracteres echappes
		$raw =~ s/_FIN_/$1/g; # caracteres echappes
		
		while($raw ne "")
		{
			if($raw =~ /^#LTAG#TITRE([1-3])#RTAG##LBLOC#([^#]*)#RBLOC#/)
			{
				creer_titre($1,$2);
				$raw = $`.$';
			}
			elsif($raw =~ /^#LTAG#CODE#RTAG##LBLOC#([^#]*)#RBLOC#/)
			{
				creer_code($1);
				$raw = $`.$';
			}		
			elsif($raw =~ /^#LTAG#ENTETE#RTAG##LBLOC#([^#]*)#RBLOC#/)
			{
				creer_EntetePage($1);
				$raw = $`.$';
				
			}
			elsif($raw =~ /^#LTAG#PIED#RTAG##LBLOC#([^#]*)#RBLOC#/)
			{
				creer_PiedPage($1);
				$raw = $`.$';
				
			}			
			elsif($raw =~ /^#LTAG#PGRAPHE#RTAG##LBLOC#([^#]*)#RBLOC#/)
			{
				creer_paragraphe($1);
				$raw = $`.$';
			}
			elsif($raw =~ /^#LTAG#LISTE#RTAG##LBLOC#([^#]*)#RBLOC#/)
			{
				my @liste = split(",",$1);
				creer_liste(@liste);
				
				$raw = $`.$';
				
			}	
			elsif($raw =~ /^#LTAG#IMAGE#RTAG##LBLOC#([0-9]+[\.]?[0-9]*,[0-9]+[\.]?[0-9]*,[^,]+,[^#]+)#RBLOC#/)
			{
				creer_image($1);
				$raw = $`.$';
			}		
			elsif($raw =~ /^#LTAG#TABLEAU#RTAG##LBLOC#([^#]*)#RBLOC#/)
			{
				my @ligne= split(":RLIGNE:,",$1);
				my @table;
				my $l = 0;
				my $c = 0;
				
				foreach my $el (@ligne)
				{
					my @col;
					$el =~ s/:LLIGNE://g;
					$el =~ s/:RLIGNE://g;
					@col = split(",",$el);
					$table[$l]= \@col;
					$c = $#col + 1;
					$l++;
				}
				if(($c > 0) && ($l > 0))
				{
					creer_tableau($l,$c,@table);
				}
				$raw = $`.$';
			}
			elsif($raw =~ /_NP_/)
			{
				creer_saut_page("");
				$raw = $';
			}	
			else{
				print "Bad format\n";
				print "^^^^^".$raw."\n";
				$raw = "";
			}
		}
	}
}


#_________________________________________
#
# Creer les styles du documents a partir
# du fichier style
#_________________________________________
sub creer_style
{
	my ($filename) = @_;
	
	open(FICHIER ,"<".$filename) or die("Ereur ouverture fichier style");
	
	$raw = "";
	
	while($ligne = <FICHIER>)
	{
		chomp($ligne);
		$raw = $raw.$ligne;
	}
	
	close(FICHIER);	
	
	$raw =~ s/\t//g;
	$raw =~ s/([\s])[\s]{1,}/$1/g;
	
	while($raw ne "")
	{
		if( $raw =~ /\[([a-zA-Z0-9]+)\]\{\{([^\}]*)\},\{([^\}]*)\},\{([^\}]*)\}\}/){
			my @type = split(",",$2);
			my @p1 = split(",",$3);
			my @p2 = split(",",$4);
			ajouter_style($1,\@type,\@p1,\@p2);
			$raw = $`.$';
		}
		else{
			$raw = "";
		}
	}
}

#_________________________________________
#
# Ajoute un style au document
#_________________________________________
sub ajouter_style
{
	my ($nom,$type,$prop1,$prop2) = @_;
	
	my %type;
	my %propa;
	my %propb;

	foreach my $e (@$type)
	{
		my @elem = split("[\'\"]:[\'\"]",$e);
		$elem[0] =~ s/[\'\"]//g;
		$elem[1] =~ s/[\'\"]//g;
		$type{$elem[0]} = $elem[1];
	}
	
	foreach my $e (@$prop1)
	{
		my @elem = split("[\'\"]:[\'\"]",$e);
		$elem[0] =~ s/[\'\"]//g;
		$elem[1] =~ s/[\'\"]//g;
		$propa{$elem[0]} = $elem[1];
	}
	foreach my $e (@$prop2)
	{
		my @elem = split("[\'\"]:[\'\"]",$e);
		$elem[0] =~ s/[\'\"]//g;
		$elem[1] =~ s/[\'\"]//g;
		$propb{$elem[0]} = $elem[1];
	}
	
	$deco->createStyle
	(
		$nom,
		%type,
		properties	=> \%propa
	);
	
	$deco->updateStyle
	(
		$nom,
		properties	=> \%propb
	);	
	
}
