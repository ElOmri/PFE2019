import * as React from 'react';

import {IUserItemProp} from '../../Modele/IUserItemProp';
import {TempState} from '../../Modele/TempState';



export class Template2 extends React.Component<IUserItemProp,TempState> {

	constructor(props: IUserItemProp, state : TempState) {
		super(props);
if(props.templatemodele==""){
		this.state = {
	
			htmlcode :
			`
	<!DOCTYPE html>
	<html lang="fr"><head>
	<meta http-equiv="content-type" content="text/html; charset=UTF-8">
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	
	<title>Sheldon COOPER - Physicien surdoué et Geek qualifié</title>
	
	<!--[if lt IE 9]>
	<script src="//html5shim.googlecode.com/svn/trunk/html5.js"></script>
	<![endif]-->
	
	<!-- Fichiers CSS -->
	<style>
	/**********
	/* Grille - 960.gs
	**********/
	.container_16 {
		max-width: 960px;
		width: 92%;
		margin: 0 auto;
	}
	.grid_6,
	.grid_8,
	.grid_10,
	.grid_16 {
		display:inline;
		float: left;
		position: relative;
		margin-left: 1%;
		margin-right: 1%;
	}
	
	.container_16 .grid_6 {
		width:35.5%;
	}
	
	.container_16 .grid_8 {
		width:48.0%;
	}
	
	.container_16 .grid_10 {
		width:60.5%;
	}
	
	.container_16 .grid_16 {
		width:98.0%;
	}
	
	.clearfix:after {
		clear: both;
		content: ' ';
		display: block;
		font-size: 0;
		line-height: 0;
		visibility: hidden;
		width: 0;
		height: 0;
	}
	
	.clearfix {
		display: inline-block;
	}
	
	* html .clearfix {
		height: 1%;
	}
	
	.clearfix {
		display: block;
	}
	
	
	
	/**********
	/* Header
	**********/
		header[role=banner] {
			background:#fff;
			border-bottom:4px solid #33a4c9;
			padding: 19px 0;
			overflow: hidden;
			
		}
		
		header figure {
			border: 3px solid #eee;
			border-radius: 78px;
			float: right;
			height: 156px;
			width: 156px;
		}
		
		header figure img {
			border-radius: 75px;
			height: 150px;
			position: relative;
			width: 150px;
		}
		
		header hgroup {
			float: left;
			padding: 42px 0;
		}
		
		header hgroup h1 {
			color:#555;
			font-family:'DroidSansBold', Arial, sans-serif;
		}
		
		header hgroup h2 {
			color:#cecece;
			font-style: italic;
			text-shadow: 1px 1px 0 rgba(0,0,0,.3);
		}
	
	/************
	/* Contact
	************/
		::-webkit-input-placeholder  { color:#fff; }
		input:-moz-placeholder { color:#fff; }
		.placeholder { color: #fff; }
	
		.contactform {
			background-color: #33a4c9;
			color: #fff;
			display: none;
		}
		
		.contactform > div {
			padding: 21px 0;
		}
		
		.contactform input[type=text], .contactform input[type=email], .contactform textarea {
			background: #2983a0;
			border: 1px solid #226d85;
			border-radius: 2px;
			color: #fff;
			max-width: 100%;
			opacity: .7;
			padding: 10px 10px 9px 10px;
			margin-top: 0;
			position: relative;
			vertical-align: top;
			width: 100%;
		}
		
		.contactform input[type=text], .contactform input[type=email] {
			margin-bottom: 21px;
		}
		
		.contactform textarea {
			font-family: 'DroidSansRegular', Arial, sans-serif;
			font-size: 14px;
			height: 148px;
		}
		
		.contactform input[type=text]:focus, .contactform input[type=email]:focus, .contactform textarea:focus {
			border: 1px solid #226d85;
			opacity: 1;
			outline: none;
			-webkit-transition: opacity 1s;
		}
		
		.contactform input[type=submit] {
			background: #ff9f39;
			border: 1px solid #ff9f39;
			border-radius: 2px;
			color: #fff;
			cursor: pointer;
			font-weight: bold;
			font-size: 14px;
			margin: 0;
			padding: 10px 10px 9px 10px;
			width: 100%;
		}
		
		.contactform input.error, .contactform textarea.error {
			border: 1px solid red;
		}
		
		.contactform label.error {
			display: none !important;
		}
	
		/***********************
		/* Message du formulaire
		***********************/
			.messageform { 
				display: none; /* Affiché via jQuery */ 
				border-radius: 2px; 
				font-size: 14px;
				font-weight: bold;
				line-height: 17px;
				padding:10px 10px 9px 10px; 
				margin: 0;
				text-align: center;
				width: 100%; 
			}
			.envoi-valid, .envoi-error { 
				display: inline-block; 
			 }
			.envoi-valid { 
				background-color: #31D869;
				border:1px solid #108D3A; 
				color: #108D3A;
			}
			.envoi-error { 
				background-color: #FF5839;
				border:1px solid #A62913; 
				color: #A62913;
			}
			.envoi-error a {
				color: #A62913;
				text-decoration: underline;
			}
			
	
	/***************
	/* Corps - Main
	***************/
		section[role=main] {
			padding: 21px 0;
			overflow: hidden;
		}
		
		section[role=main] > div {
			margin-top: 63px;
		}
		
		section[role=main] h3 {
			border-bottom: 1px solid #e1e1e1;
			padding: 0 0 20px 60px;
			margin-top: 0px;
			position: relative;
		 }
		
		section[role=main] h3:after, h3:before {
			content: ' ';
			display: block;
			height: 50px;
			left: 0;
			position: absolute;
			top: -15px;
			width: 50px;
		}
		
		section[role=main] h3:before {
			background-color: #33a4c9;
			border-radius: 25px;
		}
		
			/***************
			/* Pictos titres
			***************/
			
			section[role=main] h3:after {
				background:url('../img/pictos-titre.png') no-repeat 11px 10px;
			}
			
			section[role=main] .competences h3:after {
				background-position: 11px -34px;
			}
		
			section[role=main] .experiences h3:after {
				background-position: 11px -86px;
			}
			
			section[role=main] .formations h3:after {
				background-position: 10px -139px;
			}
			
			section[role=main] .loisirs h3:after {
				background-position: 10px -250px;
			}
			
			section[role=main] .contact h3:after {
				background-position: 11px -194px;
			}
		
		
		.experiences ul, .formations ul {
			margin-left: 60px;
		}
		
		section[role=main] h4 {
			margin-bottom: 0px;
		}
		
		section[role=main] h4 strong {
			color: #147393;
			font-weight: normal;
		}
		
		.experiences li, .formations li {
			margin-bottom: 63px;
		}
		
		.experiences li:last-child, .formations li:last-child {
			margin-bottom: 21px;
		}
		
		.experiences li p, .formations li p {
			margin-left: 30px;
		}
		
		/**********************************************
		/* Affichage des compétences sous forme de tags
		**********************************************/
		.competences .tags li {
			background-color:#33a4c9;
			border:1px solid #1b91b7;
			border-radius: 15px;
			color:#fff;
			display: inline-block;
			line-height: 13px;
			margin: 0 7px 17px 0;
			padding: 5px 10px;
		}
		
		/************************************************
		/* Affichage des compétences sous forme de barres
		*************************************************/
		.competences .barres li {
			margin-bottom: 21px;
			position: relative;
		}
		
		.competences .barres li:after {
			background: #e1e1e1;
			bottom: -4px;
			content: ' ';
			display: block;
			height: 5px;
			position: absolute;
			width:100%;
		}
		
		.competences .barres li span {
			background: #33a4c9;
			bottom: -4px;
			content: ' ';
			display: block;
			height: 5px;
			left: 0;
			position: absolute;
			width: inherit;
			z-index: 1;
		}
		
		
		/***********
		/* Pictos
		***********/
		.lieu, .dates, .phone, .mail, .site, .form, .twitter, .facebook, .dribbble, .skype {
			background: url('../img/pictos-gris.png') no-repeat;
			padding-left: 25px;
			margin: 0 10px 1px 10px;
		}
		
		.lieu {
			background-position: 0 -2px;
			color: #999;
		}
		
		.dates {
			background-position: 0 -22px;
			color: #999;
		}
		
		.phone {
			background-position: 0 -40px;
		}
		
		.mail {
			background-position: 0 -58px;
		}
		
		.site {
			background-position: 0 -80px;
		}
		
		.form {
			background-position: 0 -100px;
		}
		
		.twitter {
			background-position: 0 -122px;
		}
		
		.facebook {
			background-position: 0 -140px;
		}
		
		.dribbble {
			background-position: 0 -160px;
		}
		
		.skype {
			background-position: 0 -180px;
		}
	/* 	RESET CSS
		http://meyerweb.com/eric/tools/css/reset/ 
		v2.0 | 20110126
			 License: none (public domain)
	*/
	
	html, body, div, span, applet, object, iframe,
	h1, h2, h3, h4, h5, h6, p, blockquote, pre,
	a, abbr, acronym, address, big, cite, code,
	del, dfn, em, img, ins, kbd, q, s, samp,
	small, strike, strong, sub, sup, tt, var,
	b, u, i, center,
	dl, dt, dd, ol, ul, li,
	fieldset, form, label, legend,
	table, caption, tbody, tfoot, thead, tr, th, td,
	article, aside, canvas, details, embed, 
	figure, figcaption, footer, header, hgroup, 
	menu, nav, output, ruby, section, summary,
	time, mark, audio, video {
		margin: 0;
		padding: 0;
		border: 0;
		font-size: 100%;
		font: inherit;
		vertical-align: baseline;
	}
	/* HTML5 Reset pour les anciens navigateurs */
		article, aside, details, figcaption, figure, 
		footer, header, hgroup, menu, nav, section {
			display: block;
		}
		
	body {
		line-height: 1;
	}
	ol, ul {
		list-style: none;
	}
	blockquote, q {
		quotes: none;
	}
	blockquote:before, blockquote:after,
	q:before, q:after {
		content: '';
		content: none;
	}
	table {
		border-collapse: collapse;
		border-spacing: 0;
	}
	
	/* 	Utilisation de la propriété "box-sizing".
		Largeur d'un bloc = width + padding + border 
	*/
		* {
			-moz-box-sizing: border-box;
			box-sizing: border-box;
		}
	
	/*	Modification de la couleur et couleur de fond
		lors de la selection du contenu
	*/
		::selection {
			background: #ff9f39;
			color: #fff;
		}
		
		::-moz-selection {
			background: #ff9f39;
			color: #fff;
		}
	
	/* 	Typographie 
		- Déclaration des polices
		- Taille des polices / Interlignages
	*/
	
	/* Font-Face : déclaration de la police de caractères */
	@font-face {
			font-family: 'DroidSansRegular';
			src: url('../font/DroidSans-webfont.eot');
			src: url('../font/DroidSans-webfont.eot?#iefix') format('embedded-opentype'),
					 url('../font/DroidSans-webfont.woff') format('woff'),
					 url('../font/DroidSans-webfont.ttf') format('truetype'),
					 url('../font/DroidSans-webfont.svg#DroidSansRegular') format('svg');
			font-weight: normal;
			font-style: normal;
	
	}
	
	@font-face {
			font-family: 'DroidSansBold';
			src: url('../font/DroidSans-Bold-webfont.eot');
			src: url('../font/DroidSans-Bold-webfont.eot?#iefix') format('embedded-opentype'),
					 url('../font/DroidSans-Bold-webfont.woff') format('woff'),
					 url('../font/DroidSans-Bold-webfont.ttf') format('truetype'),
					 url('../font/DroidSans-Bold-webfont.svg#DroidSansBold') format('svg');
			font-weight: normal;
			font-style: normal;
	
	}
	
	#slider {
		/* Check vertical-rythm 
		background-image: -webkit-linear-gradient(top, #ffffff, #ffffff 96%, #000000 100%); 	
		background-size:21px 21px;
		background-position: 0 -2px; */
		background: #f5f5f5;
		color:#333;
		font-family: 'DroidSansRegular', Arial, sans-serif;
		font-size: 14px;
		line-height: 21px;
	}
	
	p, ul, h1, h2, h3, h4 {
		margin: 21px 0;
	}
	
	h1 {
		font-size: 36px;
	}
	
	h2 {
		font-size: 30px;
	}
	
	h3 {
		font-size: 24px;
		line-height: 21px;
	}
	
	h4 {
		font-size: 18px;
		line-height:21px;
	}
	
	b,strong {
		font-weight: bold;
	}
	
	em, i {
		font-style: italic;
	}
	
	/* Liens */
		a {
			color: #33a4c9;
			cursor: pointer;
			text-decoration: none;
		}
	
		a:hover, a:focus {
			color: #2c8eae;
			text-decoration: underline;
			-webkit-transition: all 1s;
		}
	
	
	
	
	</style>
	</head>
	
	<body>
	
		<!-- Header -->
		<header role="banner">
			<div class="container_16">
					<hgroup>
						<h2>user_displayName</h1>
						
						<h2>user_jobTitle - user_department</h2>
					</hgroup>
	
					<figure>
						<img src="https://cdn2.iconfinder.com/data/icons/user-interface-blue/64/User_Interface_Filled_Outline-01-512.png" alt="profilephoto">
					</figure>
			</div>
		</header>
		
		<!-- Contact -->
		<section class="contactform clearfix">
			<div class="container_16">
				<h3>Contactez-moi</h3>
				<p>Remplissez le formulaire ci-dessous afin de m'envoyer un message. Je vous répondrais dans les plus bref délai. 
				<br><em>Tous les champs sont requis.</em></p>
				<form novalidate="novalidate" method="post" action="#" name="contact" class="grid_16">
					<p class="grid_10"><textarea name="message" placeholder="Votre message" class="required"></textarea></p>
					<p class="grid_6">
						<input name="nom" placeholder="Nom - Prénom" class="required" type="text">
						<input name="email" placeholder="Adresse email" class="required" type="email">	
						<input name="envoi" value="Envoyer le message" class="required" type="submit">
						<span class="messageform"></span>
					</p>
				</form>
			</div>
		</section>
		
		<!-- Corps -->
		<section role="main" class="container_16 clearfix">
			<div class="grid_16">
				<!-- A propos -->
				<div class="grid_8 apropos">
					<h3>A propos</h3>
					<p>Cette section vous sert de présentation.</p>
					<p>user_aboutme</p>
				</div>
				
				<!-- Compétences -->
				<div class="grid_8 competences">
					<h3>Compétences</h3>
					<ul class="barres">
						user_skills
					</ul>
				</div>
			</div>
			
				<!-- Expériences -->
				<div class="grid_16 experiences">
					<h3>Expériences</h3>
					<ul>
						<li>
							<h4><strong>Nom du poste</strong> chez nom de l'employeur</h4>
							<span class="lieu">Lieu</span>
							<span class="dates">Dates</span>
							<p>Projet : Descreption du projet </p>
						</li>
						user_pastProjects
					</ul>
				</div>
			
				<!-- Formations -->
				<div class="grid_16 formations">
					<h3>Formations</h3>
					<ul>
						<li>
							<h4><strong>Nom de la formation / diplôme</strong> à nom de l'école</h4>
							<span class="lieu">Lieu</span>
							<span class="dates">Période</span>
							<p>Clarification de la formation (les abréviations ne sont pas connues de tout le monde).</p>
						</li>
						user_schools
				</div>
			
				<!-- Loisirs -->
				<div class="grid_8 loisirs">
					<h3>Loisirs</h3>
				 user_interests
				</div>
			
				<!-- Contact -->
				<div class="grid_8 contact">
					<h3>Contact</h3>
					<p>Si mon profil vous intéresse, n'hésitez pas à me contacter .</p>
					<ul>
						<li class="lieu">  </li>
						<li class="phone">ID : user_id  </li>
						<li class="mail"><a href="mailto:mon.adresse@email.fr">  </a></li>
					</ul>
				</div>
		</section>
	
	
	
	</body></html>
`
		};
	}
	else
	{
		this.state={
			htmlcode:this.props.templatemodele,
		}
	}
}
	image;
	private CreateSkills() :string {
    let table = ''
    for (let i = 0; i <this.props.skills.length; i++) {
				
      table+=(`<li data-skills="80">`+this.props.skills[i]+`<span style={{width: '80'}}></span></li>`)
    }
    return table
	}
	
	private CreateSchools() :string {
    let table = ''

    for (let i = 0; i <this.props.schools.length; i++) {
      
			table+=(`<li data-skills="80"><strong>`+this.props.schools[i]+`</strong><span style={{width: '80%'}}></span></li>`)
    }
    return table
	}
	private CreatePastProjects() :string{
		let table = ''
			
		for (let i = 0; i <this.props.pastProjects.length; i++) {
			let	verif=true;

			if(this.props.projects){
				this.props.projects.value.forEach(element => {
					if(element.fields.Title==this.props.pastProjects[i])
					{
					
						verif=false;
						table+=(`
						<li><strong>`+this.props.pastProjects[i]+" : " +element.fields.Description+`</strong> </li>`)
					}
	
				}
				
				);
			}
			if(verif)
			{
				table+=(`
				<li><strong>`+this.props.pastProjects[i]+`</strong> </li>`)
			}
	
	
		}
		return table
		}

	

	private Createresponsibilities() :string{
    let table = ''

    for (let i = 0; i <this.props.responsibilities.length; i++) {
      
			table+=(`
			<li><strong>`+this.props.responsibilities[i]+`</strong> </li>`)
    }
    return table
	}

	private CreateInterests() :string{
    let table = ''

    for (let i = 0; i <this.props.interests.length; i++) {
      
      table+=(`<p><strong>`+this.props.interests[i]+`</strong> </p>`)
    }
    return table
	}

	private urlimage() : string
	{
		return 'https://eur.delve.office.com/mt/v3/people/profileimage?userId='+this.props.userPrincipalName+'&size=L'
	}
	


	


private TemplateRenderAF=function()  {

	var html= this.state.htmlcode;
	var skills= this.CreateSkills();
	var schools= this.CreateSchools();
	var pastProjects= this.CreatePastProjects();
	var interests= this.CreateInterests();
	var responsibilities= this.Createresponsibilities();
	var urlphoto = this.urlimage();
	var str =html.replace(/user_mail/g,this.props.mail);
	var str2=str.replace(/user_city/g,this.props.city);
	var str3 =str2.replace(/user_id/g,this.props.id);
	var str4 =str3.replace(/user_aboutme/g,this.props.aboutMe);
	var str5 =str4.replace(/user_department/g,this.props.department);
	var str6 =str5.replace(/user_jobTitle/g,this.props.jobTitle);
	var str7 =str6.replace(/user_displayName/g,"");
	var str8 =str7.replace(/user_skills/g,skills);
	var str9 =str8.replace(/user_pastProjects/g,pastProjects);
	var str10 =str9.replace(/user_schools/g,schools);
	var str11 =str10.replace(/user_interests/g,interests);
	var str12 =str11.replace(/urlphoto/g,urlphoto);
	var finalstr=str12.replace(/user_responsibilities/g,responsibilities);
	return(
<div>
<div id={this.props.id}  style={{
					backgroundColor: '#f5f5f5',
				 
				}} >
<span style={{width:"100%"}} dangerouslySetInnerHTML={{__html:finalstr}} />
</div>
</div>
	)

}
    render() : JSX.Element {
    return (
			
<div>		
{this.TemplateRenderAF()}

</div>	


    	    );
  }
 

}






