<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
	xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" version="1.0"
	exclude-result-prefixes="w r">

	<xsl:output indent="no" encoding="UTF-8" version="1.0"/>
	
	
	
	<!-- 
	Deuxième passage pour activer les liens hypertextes inactifs.
	Le premier passage dans document.xml ajoute des balises <w:hyperlink> aux liens inactifs, et leur
	attribue un @r:id aléatoire.
	Ce deuxième passage ajoute dans document.xml.rels, des balises <Relationship> correspondants aux
	balises <w:hyperlink> dans document.xml qui n'ont pas de correspondance ici.
	-->
    
    
    
    <xsl:param name="file"/>
    
    <xsl:variable name="pathfile">
        <xsl:value-of select="concat('/word/', $file, '.xml')"/>
    </xsl:variable>



	<!-- 
	Le script Python passe dans $tempdir le chemin absolu vers le répertoire temporaire où se trouve le fichier
	XML source.
	La lib Python lxml ne gère pas bien la fonction XPath document() et n'accepte que des chemins
	relatifs à la feuille XSLT, ou des chemins absolus (2ème argument de la fonction document() non géré).
	-->
	<xsl:param name="tempdir"/>
	<xsl:variable name="tempdirSlash" select="translate($tempdir, '\', '/')"/>
	
	
	
	<xsl:variable name="root" select="/"/>



	<!-- Copie par défaut. -->
	<xsl:template match="node() | @*">
		<xsl:copy>
			<xsl:apply-templates select="node() | @*"/>
		</xsl:copy>
	</xsl:template>



	<xsl:template match="/*[local-name()='Relationships']">
		<xsl:copy>			
			<!-- apply-templates pour les éléments <Relationship> déjà présents dans le document courant -->
			<xsl:apply-templates select="node() | @*"/>
			<!-- apply-templates pour les éléments <w:hyperlink> de document.xml, dont on va tester l'@r:id  -->
		    <xsl:apply-templates select="document(concat($tempdirSlash, $pathfile))//w:hyperlink/@r:id"/>
		</xsl:copy>
	</xsl:template>



	<!--
	Détermine si, dans le fichier document.xml préalablement traité, il existe des éléments
	<w:hyperlink> nouvellement créés, dont l'@r:id n'a pas d'équivalent dans les éléments
	<Relationship> du document courant.
	Dans ce cas crée un élément <Relationship> avec l'@Id et le lien correspondant.
	-->
	<xsl:template match="//w:hyperlink/@r:id">
		<xsl:choose>
			<!--
			La racine courante de cette template, c'est celle du fichier document.xml
			Pour atteindre des éléments du fichier document.xml.rels, le document source de cette transformation, il
			faut partir de la bonne racine.
			Qui est stockée dans $root pour être accessible ici dans cette template.
			-->
			<!-- S'il y a un élément <Relationship> correspondant : rien -->
			<xsl:when
				test=". = $root/*[local-name()='Relationships']/*[local-name()='Relationship']/@Id"/>
			<!-- Sinon : on en crée un avec le bon @Id. -->
			<xsl:otherwise>
				<Relationship Id="{.}"
					Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink">
					<xsl:attribute name="Target">
						<xsl:choose>
							<xsl:when test="starts-with(..//w:t/text(), 'www.')">
								<xsl:value-of select="concat('http://', ..//w:t/text())"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="..//w:t/text()"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				    <xsl:attribute name="TargetMode">
				        <xsl:value-of select="'External'"/>
				    </xsl:attribute>
				</Relationship>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>
