<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
	xmlns:o="urn:schemas-microsoft-com:office:office"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
	xmlns:v="urn:schemas-microsoft-com:vml"
	xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	xmlns:w10="urn:schemas-microsoft-com:office:word"
	xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" version="1.0">

	<xsl:output indent="no" encoding="UTF-8" version="1.0"/>
	
	
	
	<xsl:variable name="root" select="/"/>
	
	
	
    <!-- Copie par défaut. -->
	<xsl:template match="node() | @*">
		<xsl:copy>
			<xsl:apply-templates select="node() | @*"/>
		</xsl:copy>
	</xsl:template>



	<!--
	Sélectionne les text run (w:r) dont le texte (w:t) contient 'http' ou 'https' ou 'ftp' ou 'sftp',
	et qui ne sont pas encadrés dans des balises <w:hyperlink> (donc pas activés).
	Les encadre dans un élément <w:hyperlink> dont l'id est généré aléatoirement.
	
	Cette template ne gère que les liens qui se trouvent seuls dans des balises <w:t>
	S'ils contiennent du texte supplémentaire avant ou après (avec des espaces ' '), le lien n'est pas traité.
	
	Je ne sais pas pourquoi MS Word isole les liens hypertextes dans des balise <w:t> ni si c'est
	systématique.
	Je l'ai constaté de manière empirique.
	S'il avait fallu traiter des liens mélangés à du texte, ça aurait été beaucoup plus complexe.
	La doc explique ce que veulent dire les balises, pas la manière dont MS Word les utilise...
	-->
	<xsl:template
		match="//w:r[not(parent::w:hyperlink)][w:t[starts-with(text(), 'http://') or starts-with(text(), 'https://') or
		starts-with(text(), 'ftp://') or starts-with(text(), 'sftp://') or starts-with(text(), 'www.')][not(contains(text(), ' '))]]">
		<w:hyperlink r:id="{generate-id()}">
			<xsl:copy>
				<xsl:copy-of select="@*"/>
				<xsl:apply-templates/>
			</xsl:copy>
		</w:hyperlink>
	</xsl:template>
</xsl:stylesheet>
