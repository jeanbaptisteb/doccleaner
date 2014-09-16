<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/" version="1.0" exclude-result-prefixes="w r dc dcterms">

    <xsl:output indent="no" encoding="UTF-8" version="1.0"/>



    <!-- 
	Deuxième passage pour activer les liens hypertextes inactifs.
	Le premier passage dans document.xml ajoute des balises <w:hyperlink> aux liens inactifs, et leur
	attribue un @r:id aléatoire.
	Ce deuxième passage ajoute dans document.xml.rels, des balises <Relationship> correspondants aux
	balises <w:hyperlink> dans document.xml qui n'ont pas de correspondance ici.
	-->


    <!-- 
	Le script Python passe dans $tempdir le chemin absolu vers le répertoire temporaire où se trouve le fichier
	XML source.
	La lib Python lxml ne gère pas bien la fonction XPath document() et n'accepte que des chemins
	relatifs à la feuille XSLT, ou des chemins absolus (2ème argument de la fonction document() non géré).
	-->
    <xsl:param name="tempdir"/>
    <xsl:variable name="tempdirSlash" select="translate($tempdir, '\', '/')"/>
    
    
    
    <xsl:variable name="root" select="/"/>
    
    
    <!--
    <xsl:variable name="stylesdocument" select="document(concat($tempdirSlash,'/word/styles.xml'))"/>
    <xsl:variable name="commentsdocument" select="document(concat($tempdirSlash,'/word/comments.xml'))"/>
    <xsl:variable name="coredocument" select="document(concat($tempdirSlash,'/docProps/core.xml'))"/>
    -->
    <xsl:variable name="stylesdocument" select="document('styles.xml', /)"/>
    <xsl:variable name="commentsdocument" select="document('comments.xml', /)"/>
    <xsl:variable name="coredocument" select="document('../docProps/core.xml', /)"/>
    
    

    <!-- 
    On teste si le fichier comments.xml existe ou non, pour savoir si on modifie le fichier
    existant ou si on en crée un nouveau depuis zéro.
    -->
    <xsl:variable name="commentsFileExists" select="boolean($commentsdocument)"/>



    <xsl:template match="/">
        <xsl:choose>
            <xsl:when test="$commentsFileExists = false()">
                <xsl:call-template name="noComments"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:call-template name="comments"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <!-- 
    ========================================================
    Création d'un fichier comments.xml quand il n'existe pas
    ========================================================
    -->

    <!-- Pour chaque commentaire dans document.xml, on ajoute une balise w:comment dans le fichier -->
    <xsl:template name="noComments">
        <w:comments xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
            
            <!-- On cherche tous les commentaires du fichier document.xml -->
            <xsl:apply-templates select="descendant::w:commentReference"/>
            <!--
            <xsl:apply-templates select="descendant::w:commentReference
                | document('footnotes.xml', /)/descendant::w:commentReference
                | document('endnotes.xml', /)/descendant::w:commentReference"/>
             -->
            
        </w:comments>
    </xsl:template>
    
    
    
    <!-- 
    ====================================================
    Complétion d'un fichier comments.xml qui existe déjà
    ====================================================
    -->

    <xsl:template name="comments">
        <xsl:apply-templates select="$commentsdocument" mode="comments"/>
    </xsl:template>
    
    
    
    <!--
    On commence par recopier tout le fichier, puis pour chaque commentaire dans document.xml
    qui n'a pas d'équivalent dans comment.xml, on ajoute une balise w:comment.
    -->
    <xsl:template match="w:comments" mode="comments">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*" mode="comments"/>
            
            <!--
            <SUITE/>
            -->
            
            <xsl:apply-templates select="$root/descendant::w:commentReference[@w:id !=
                current()/descendant::w:comment/@w:id]"/>
            
            <!-- 
            Appliquer également sur footnotes.xml et endnotes.xml en changeant $root
            -->
        </xsl:copy>
    </xsl:template>



    <!-- Recopiage du fichier comment.xml -->
    <xsl:template match="node() | @*" mode="comments">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*" mode="comments"/>
        </xsl:copy>
    </xsl:template>
    
    
    
    
    <!-- 
    ===============================================================
    Template commune à la création ou la complétion de comments.xml
    ===============================================================
    -->
    
    <!--
    La template crée une balise w:comment dans comments.xml lorsqu'elle est appelée.
    Elle est appelée lorsqu'il y a un commentaire sans équivalence dans document.xml,
    endnotes.xml ou footnotes.xml
    -->
    <xsl:template match="w:commentReference" name="commentReference">
        <w:comment w:id="{@w:id}" w:author="{$coredocument//dc:creator}"
            w:date="{$coredocument//dcterms:modified}" w:initials="r">
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="{$stylesdocument//w:style[w:name/@w:val='annotation text']/@w:styleId}"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="{$stylesdocument//w:style[w:name/@w:val='annotation reference']/@w:styleId}"/>
                    </w:rPr>
                    <w:annotationRef/>
                </w:r>
                <w:r>
                    <w:t>
                        <xsl:choose>
                            <xsl:when test="substring(@w:id, string-length(@w:id)) = 1">
                                <xsl:text>Paragraphe sans majuscule ni ponctuation finale.</xsl:text>
                            </xsl:when>
                            <xsl:when test="substring(@w:id, string-length(@w:id)) = 2">
                                <xsl:text>Paragraphe sans ponctuation finale.</xsl:text>
                            </xsl:when>
                            <xsl:when test="substring(@w:id, string-length(@w:id)) = 3">
                                <xsl:text>Paragraphe sans majuscule.</xsl:text>
                            </xsl:when>
                        </xsl:choose>
                    </w:t>
                </w:r>
            </w:p>
        </w:comment>
    </xsl:template>



    <!-- Pas d'interférence du document source -->
    <xsl:template match="node() | text()"/>
</xsl:stylesheet>
