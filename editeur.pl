#!/bin/perl

#________________________________________________
# Auteur: Rija ZAFIAMY
# Date : fevrier 2015
# Generateur de document openoffice
#________________________________________________

use OpenOffice::OODoc;

if((not defined $ARGV[0]) &&(not defined $ARGV[1] ))
{
	print "Syntax error : $0 <-f|model> <text_to_format|style>\n";
	exit 1;
}

#________________________________________________
# Variables globales
#________________________________________________
my @contenu;
my @titre1;
my @titre3;
my $t1;
my $t2;
my $t3;
my $meta;
my $texte;
my $deco;
my $master;
my $layout;
my $archive;
my $countTab;
my $countImg;

if($ARGV[0] eq "-f")
{
	formater($ARGV[1]);
}
else
{
	($t1,$t2,$t3,$countTab,$countImg)	= (0,1,0,0,0);
	
	@titre1		= ("I","II","III","IV","V","VI","VII","VIII",
					"IX","X","XI","XII","XIII","XIV","XV",
					"XVI","XVII","XVIII","XIX","XX","XXI",
					"XXII","XXIII","XXIV","XXV","XXVI");

	@titre3		= ("a","b","c","d","e","f","g","h","i","j","k",
					"l","m","n","o","p","q","r","s","t","u","v",
					"w","y","z");
				
	# 1) Creee un document ODT vide
	$archive 							= creerDoc($ARGV[0]);
	
	# 2) Initialise l'objet oodoc
	($meta,$texte,$deco) 				= initDoc($archive);
	
	# 3) Definit le titre du document
	$meta->title('Rapport');
	
	# 4) Creee un style de page
	$layout = $deco->pageLayout("Standard");
	$master = $deco->createMasterPage("StylePages",layout=>$layout);
	
	# 5) Creee tous les styles du document a partir de style.sty
	creer_styles($ARGV[1]);
	
	# 6) Effectue un pre-traitement sur le fichier
	my $raw = preprocesseur($ARGV[0]);
	
	# 7) Traite le document POF
	automate($raw);
	
	# 8) Met a jour le fichier ODT
	$archive->save;
}

# --------------------------------------
# Creation d'un document ODT vide
#---------------------------------------
sub creerDoc
{
	my ($nom_fichier) = @_;
	
	my $doc = ooDocument
	(
		file 	=> $nom_fichier.".odt",
		create 	=> 'text',
		member	=> 'content'
	);

	$doc->save;
	
	my $archive 	= ooFile($nom_fichier.".odt");
	
	return $archive;
}
# --------------------------------------
# Instance objet ODoc
# return:
#	- meta
#	- texte
#	- deco
#---------------------------------------
sub initDoc
{
	my ($fic) = @_;

	# instance objet Meta
	my $meta 		= ooMeta
	(
		archive => $fic
	);
	# instance objet Document pour content
	my $texte 		= ooDocument
	(
		archive => $fic,
		member 	=> 'content'
	);
	# instance objet Document pour les styles
	my $deco = ooDocument
	(
		archive => $fic,
		member 	=> 'styles'
	);
	
	return ($meta,$texte,$deco);
}

#_________________________________________
#
# Insere un titre de niveau 1-3 dans le document
#_________________________________________
sub creer_titre
{
	my ($niveau,$titre) = @_;
	$titre =~ s/_TAB_/\t/g;
	
	if(1==$niveau)
	{
		$t2 = 1;
		$t3 = 0;
		$titre = $titre1[$t1]."- ".$titre;
		$t1++;
	}
	elsif(2==$niveau)
	{
		$t3 = 0;
		$titre = $t2.". ".$titre;
		$t2++;
	}
	else
	{
		$titre = $titre3[$t3].") ".$titre;
		$t3++;
	}
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
	
	my ($tabtitle,$l,$c,@tab) = @_;
		
	my $feuille = $texte->appendTable('Table'.$countTab, $l, $c);

	my $i = 0;
	
	foreach my $li (@tab)
	{
		my $j = 0;
		
		foreach my $col (@$li)
		{
			$col =~ s/_TAB_//g;
			$cell = $texte->getTableCell($feuille,$i,$j);
			
			$texte->appendParagraph
				(
					text => $col,
					style => "StyleTableau",
					attachment => $cell
				);
				
			push(@contenu,$col);
			$j++;
		}
		$i++;
	}
	
	$countTab++;
	
	$texte->appendParagraph
	(
	 	style 	=> 'StyleLegende',
		text  	=> 'Table '.$countTab."  : ".$tabtitle
	);
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
			$parag = $texte->appendParagraph
			(
				style => 'StyleListe',
				text  => $item
			);
			ajouter_hypertexte($parag,$item);
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
	
	
	my $fig = $countImg + 1;
	
	my $p = $texte->appendParagraph
	(
	 	style 	=> 'StyleLegende',
		text  	=> "Figure ".$fig.": ".$param[2]
	);
	
	$texte->insertImageElement
                (
					"Image".$countImg,
					title    => $param[2],
					description     => $param[2],
					after      	=> $p,
					legend => $p,
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
		$parag = $texte->appendParagraph
		(
			style 	=> 'StyleParagraphe',
			text  	=> $p
		);
		push(@contenu,$p);
		ajouter_hypertexte($parag,$p);
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
	
	$texte->appendParagraph
	(
		style 	=> 'StyleCode',
		text  	=> ""
	);
	foreach my $c (@lignes)
	{
		$texte->appendParagraph
		(
			style 	=> 'StyleCode',
			text  	=> $c
		);
		push(@contenu,$c);
	}
	$texte->appendParagraph
	(
		style 	=> 'StyleCode',
		text  	=> ""
	);
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
# Creer des liens hypertextes
#_________________________________________
sub ajouter_hypertexte
{

	my ($p,$t) = @_;
	
	my $tmp = $t;
	
	while($tmp ne "")
	{	
		if($tmp=~ /[\s]*(http[\S]+[^.,:\s]).*/){
			$texte->setHyperlink($p,$1,$1);
			$tmp = $';
		}
		else
		{
			$tmp = "";
		}
		
	}
}
#_________________________________________
# Formate un document au format POF et 
# l'affiche a la sortie standard.
#_________________________________________
sub formater
{
	my ($nom_fichier) = @_;
	
	open(FICHIER ,"<".$nom_fichier) 
	or 
	die("Ereur ouverture fichier $nom_fichier");
	
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
#_________________________________________
# Effectue un pre-traitement sur le 
# fichier source avant de le convertir
# en fichier ODT.
#_________________________________________
sub preprocesseur
{
	my($fichier) = @_;
	
	my $raw = "";
	
	open(FICHIER ,"<".$fichier) or die("Ereur ouverture fichier modele");
	
	while($ligne = <FICHIER>)
	{
		$ligne =~ s/^%%.*//g; #supprime les commentaires
		$ligne =~ s/[\t\n\0\r]*$//g; #supprime les caractères de saut en fin de ligne
		chomp($ligne);
		$raw = $raw.$ligne;
		
	}
	
	close(FICHIER);	
	
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
	
	$raw =~ s/:[\s]*\{/#LBLOC#/g;
	$raw =~ s/([^\\]{1})\}/$1#RBLOC#/g;
	
	$raw =~ s/\\(.)/$1/g; # caracteres echappes
	$raw =~ s/_DEBUT_//g; # caracteres echappes
	$raw =~ s/_FIN_/$1/g; # caracteres echappes
	
	return $raw;
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
	
	my ($raw) = @_;
	
	
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
			
			my $isTitle = 1;
			my $tabTitle = "";
			
			foreach my $el (@ligne)
			{
				my @col;
				$el =~ s/:LLIGNE://g;
				$el =~ s/:RLIGNE://g;
				
				@col = split(",",$el);
				
				if($isTitle)
				{
					$tabTitle = $el;
					$isTitle = 0;
				
				}
				else
				{
					$table[$l]= \@col;
					$c = $#col + 1;
					$l++;
				}
			}
			if(($c > 0) && ($l > 0))
			{
				creer_tableau($tabTitle,$l,$c,@table);
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


#_________________________________________
#
# Creer les styles du documents a partir
# du fichier style
#_________________________________________
sub creer_styles
{
	my ($filename) = @_;
	
	open(FICHIER ,"<".$filename) or die("Ereur ouverture fichier style");
	
	my $raw = "";
	
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
