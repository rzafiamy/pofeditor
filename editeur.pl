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
my @tableMatiere;
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
	
	# 3) Definit les infos du document par defaut
	updateMetaInfo($meta,"REPORT","POF EDITOR","Rija ZAFIAMY","fr-FR");
	
	# 4) Creee un style de page
	$layout = $deco->pageLayout("Standard");
	$master = $deco->createMasterPage("StylePages",layout=>$layout);
	
	# 5) Creee tous les styles du document a partir de style.sty
	parse_styles($ARGV[1]);
	
	# 6) Effectue un pre-traitement sur le fichier
	my $raw = preprocesseur($ARGV[0]);
	
	# 7) Traite le document POF
	automate($raw);
	
	# foreach my $e (@tableMatiere)
	# {

		# foreach my $g (keys %$e)
		# {
			# if($g=~"3")
			# {
				# $$e{$g}= remplir_Paragrahe($$e{$g},".");
				# creer_paragraphe($$e{$g},"StyleTableMatiereTitre3");
			# }
			# elsif($g=~"2")
			# {
				# $$e{$g}= remplir_Paragrahe($$e{$g},".");
				# creer_paragraphe($$e{$g},"StyleTableMatiereTitre2");
			# }
			# elsif($g=~"1")
			# {
				# $$e{$g}= remplir_Paragrahe($$e{$g},".");
				# creer_paragraphe($$e{$g},"StyleTableMatiereTitre1");
			# }	
		# }
	# }
	
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
# Set meta information
# return:
#	- meta
#---------------------------------------
sub updateMetaInfo
{
	my ($meta,$title,$subject,$author,$lang)	= @_;

	$meta->title($title);
	$meta->subject($subject);
	$meta->creator($author);
	$meta->initial_creator($author);
	$meta->language($lang);
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
	my ($niveau,$ordre,$titre) = @_;
	$titre =~ s/_TAB_/\t/g;
	
	if(1==$ordre)
	{
		if(1==$niveau)
		{
			$t2 = 1;
			$t3 = 0;
			$titre = "Chapter ".$titre1[$t1].": ".$titre;
			$t1++;
		}
		elsif(2==$niveau)
		{	
			$t2 = 1;
			$titre = uc($titre3[$t3]).") ".$titre;
			$t3++;
		}
		else
		{
			$titre = $t2.". ".$titre;
			$t2++;
		}	
		
		my %hash  = ("$niveau" => "$titre");
	
		push(@tableMatiere,\%hash);
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
		
	my $feuille = $texte->appendTable('Tableau'.$countTab, $l, $c,"table:style-name" => 'StyleTab'.$countTab);

	my $i = 0;
	
	foreach my $li (@tab)
	{
		my $j = 0;
		
		foreach my $col (@$li)
		{
			$col =~ s/_TAB_//g;
			$col =~ s/_DIESE_/#/g;
			$cell = $texte->getTableCell($feuille,$i,$j);
			
			$texte->cellStyle($cell,"StyleCell");
			
			my $styleTab	= "StyleTableau$j";
			
			if(0==$i)
			{
				$styleTab	= "StyleTableauTete"
			}
			
			my @parag = split("_NL_",$col);
			
			foreach my $e (@parag)
			{
				my $para = $texte->appendParagraph
						(
							text => $e,
							style => $styleTab,
							attachment => $cell
						);
						
				ajouter_hypertexte($para,$e);
				push(@contenu,$col);
			}
			
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
	
	$texte->createImageElement
                (
					"Image".$countImg,
					title    => $param[2],
					description     => $param[2],
					attachment      	=> $p,
					legend => $p,
					import          => $param[3],
					size 			=> $param[0]." cm,".$param[1]." cm",
					style			=> "StyleImage"
                );
	
	creer_paragraphe("");
	
	$countImg++;
}

#_________________________________________
#
# Insere un paragraphe dans le document
#_________________________________________
sub creer_paragraphe
{
	my $paragraphe		= shift;
	my $style			= shift;
	
	if(not defined $style)
	{
		$style = "StyleParagraphe";
	
	}
	
	$paragraphe =~ s/_TAB_//g;
	$paragraphe =~ s/_DIESE_/#/g;
	$paragraphe = ajouter_tabulation($paragraphe);
	my @parag = split("_NL_",$paragraphe);
	

	foreach my $p (@parag)
	{
		my $para = $texte->appendParagraph
		(
			style 	=> $style,
			text  	=> $p
		);
		push(@contenu,$p);
		ajouter_hypertexte($para,$p);
	}
	
	my $ret;
		
	$ret = $texte->appendParagraph
	(
	 	style 	=> $style,
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
	my ($num,$code) = @_;
	
	$code =~ s/_TAB_/\t/g;
	$code =~ s/_DIESE_/#/g;
	my @lignes = split("_NL_",$code);
	
	$texte->appendParagraph
	(
		style 	=> "StyleCode$num",
		text  	=> ""
	);
	foreach my $c (@lignes)
	{
		$texte->appendParagraph
		(
			style 	=> "StyleCode$num",
			text  	=> $c
		);
		push(@contenu,$c);
	}
	$texte->appendParagraph
	(
		style 	=> "StyleCode$num",
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
# Creer un saut de page
#_________________________________________
sub creer_saut_page
{
	
	my ($param) = @_;
		
	my $p = creer_paragraphe("");
	$texte->setPageBreak($p,style=> "xyz",page=>"StylePages");
}

#_________________________________________
#
# Creer un saut de ligne
#_________________________________________
sub creer_saut_ligne
{
	
	my ($param) = @_;
		
	$texte->appendParagraph
	(
		style 	=> 'StyleParagraphe',
		text  	=> ""
	);
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
			$texte->setHyperlink(
				$p,
				$1,
				$1,
				properties => {'style' => "StyleHyperlink"});
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
	
	while(my $ligne = <FICHIER>)
	{
		$ligne =~ s/^[\s]*%%.*//g; 		#supprime les commentaires
		$ligne =~ s/[\t\n\r\0]*$//g; 	#supprime les caracteres de saut
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
	$raw =~ s/([^\\]{1})\]([\s]*):([\s]*)/$1#RTAG#:/g; #forme ]:
	
	$raw =~ s/([\s]*):([\s]*)\{/#LBLOC#/g;
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
		if($raw =~ /^#LTAG#TITRE([1-3])(\*){0,1}#RTAG##LBLOC#([^#]*)#RBLOC#/)
		{
			if($2 eq "*") # liste non ordonnnee
			{
				creer_titre($1,0,$3);
			}
			else
			{
				creer_titre($1,1,$3);
			}
			
			$raw = $`.$';
		}
		elsif($raw =~ /^#LTAG#CODE([1-3]){0,1}#RTAG##LBLOC#([^#]*)#RBLOC#/)
		{
			my $num = "";
			my $niveau = $1;
			my $code = $2;
			
			$raw = $`.$';
			
			if($niveau =~/^[1-3]$/)
			{
				$num = $niveau;
			}
			creer_code($num,$code);
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
		elsif($raw =~ /^#LTAG#DOCINFO#RTAG##LBLOC#([^#]+)#RBLOC#/)
		{
			$raw = $`.$';
			
			my @infos	= split(",",$1);
			
			foreach my $e (@infos)
			{
				my $author;
				my $title;
				my $subject;
				my $lang;
				
				if($e =~ /author:(.+)$/)
				{	
					$author	= $1;
				}
				if($e =~ /title:(.+)$/)
				{	
					$title	= $1
				
				}
				if($e =~ /subject:(.+)$/)
				{	
					$subject	= $1
				
				}
				if($e =~ /lang:(.+)$/)
				{	
					$lang	= $1
				
				}
				
				updateMetaInfo($meta,$title,$subject,$author,$lang);
			
			}
			
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
				
				@col = split(";;",$el);
				
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
		elsif($raw =~ /^_NP_/)
		{
			creer_saut_page("");
			$raw = $';
		}
		elsif($raw =~ /^_NL_/)
		{
			creer_saut_ligne("");
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
# Recupere chaque section du style 
# (propriete,nom,type,...)
#_________________________________________
sub splitStyle
{

	my ($t,$prop1,$prop2) = @_;
	
	my %type;
	my %propa;
	my %propb;
			
	foreach my $e (@$t)
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
	
	return (\%type,\%propa,\%propb);
}
#_________________________________________
#
# Creer les styles du documents a partir
# du fichier style
#_________________________________________
sub parse_styles
{
	my ($filename) = @_;
	
	open(FICHIER ,"<".$filename) or die("Ereur ouverture fichier style");
	
	my $raw = "";
	
	while($ligne = <FICHIER>)
	{
		$ligne =~ s/^[\s]*%%.+//g; #supprime les commentaires
		chomp($ligne);
		$raw = $raw.$ligne;
	}
	
	close(FICHIER);	
	
	$raw =~ s/\t//g;
	$raw =~ s/([\s])[\s]{1,}/$1/g;
	
	while($raw ne "")
	{
		if( $raw =~ /\[([a-zA-Z0-9]+)\]\{\{([^\}]*)\},\{([^\}]*)\},\{([^\}]*)\}\}/)
		{
			my @t  = split(",",$2);
			my @pa = split(",",$3);
			my @pb = split(",",$4);
			
			my ($type,$p1,$p2) = splitStyle(\@t,\@pa,\@pb);
			
			creer_style($1,$type,$p1,$p2);
			
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
sub creer_style
{
	my ($nom,$type,$propa,$propb) = @_;
	
		$deco->createStyle
		(
			$nom,
			%$type,
			properties	=> $propa
		);
		
		$deco->updateStyle
		(
			$nom,
			properties	=> $propb
		);	
	
}

#_________________________________________
#
# Remplir delimiteur
#_________________________________________
sub remplir_Paragrahe
{
	my $pgraphe 	= shift;
	my $motif		= shift;
	
	my $taille	= length($pgraphe);
	my $reste		= 60 - $taille;
	
	if($reste > 0)
	{
		for(my $i=0;$i<$reste;$i++)
		{
			$pgraphe .= $motif;
		}
	}
	
	return $pgraphe;
}
