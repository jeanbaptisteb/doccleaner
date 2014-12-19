<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
    version="1.0" exclude-result-prefixes="w r dc dcterms">

    <xsl:output indent="no" encoding="UTF-8" version="1.0"/>



    <!-- 
	Le script Python passe dans $tempdir le chemin absolu vers le répertoire temporaire où se trouve le fichier
	XML source.
	La lib Python lxml ne gère pas bien la fonction XPath document() et n'accepte que des chemins
	relatifs à la feuille XSLT, ou des chemins absolus (2ème argument de la fonction document() non géré).
	-->
    <xsl:param name="tempdir"/>
    <xsl:variable name="tempdirSlash" select="translate($tempdir, '\', '/')"/>



    <xsl:variable name="root" select="/"/>


    
    <xsl:variable name="stylesdocument" select="document(concat($tempdirSlash,'/word/styles.xml'))"/>
    <xsl:variable name="commentsdocument"
        select="document(concat($tempdirSlash,'/word/comments.xml'))"/>
    <xsl:variable name="coredocument" select="document(concat($tempdirSlash,'/docProps/core.xml'))"/>
    <!--
    <xsl:variable name="stylesdocument" select="document('styles.xml', /)"/>
    <xsl:variable name="commentsdocument" select="document('comments.xml', /)"/>
    <xsl:variable name="coredocument" select="document('docProps/core.xml', /)"/>
-->


    <xsl:key name="comment" match="//w:comment" use="ancestor::*"/>



    <xsl:variable name="commentsId">
        <xsl:call-template name="getCommentsId">
            <xsl:with-param name="count" select="count($commentsdocument//w:comment)"/>
            <xsl:with-param name="index" select="1"/>
            <xsl:with-param name="ids" select="''"/>
        </xsl:call-template>
    </xsl:variable>




    <xsl:template name="getCommentsId">
        <xsl:param name="count" select="count($commentsdocument//w:comment)"/>
        <xsl:param name="index" select="1"/>
        <xsl:param name="ids" select="''"/>
        <xsl:choose>
            <xsl:when test="$index &gt; $count">
                <xsl:value-of select="$ids"/>
            </xsl:when>
            <xsl:otherwise>
                <!-- 
                Le for-each n'itère que sur le noeud document du fichier comments.xml
                
                En fait il ne sert qu'à changer le focus de l'élément courant sur comments.xml pour
                que l'appel de key() s'effectue sur le bon document.
                -->
                <xsl:for-each select="$commentsdocument">
                    <xsl:call-template name="getCommentsId">
                        <xsl:with-param name="count" select="$count"/>
                        <xsl:with-param name="index" select="$index +1"/>
                        <xsl:with-param name="ids"
                            select="concat($ids, '*', key('comment',
                        /)[$index]/@w:id,
                        '*')"
                        />
                    </xsl:call-template>
                </xsl:for-each>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <!--
    On teste si le fichier comments.xml existe ou non, pour savoir si on modifie le fichier
    existant ou si on en crée un nouveau depuis zéro.
    NE MARCHE PAS. DOIT ETRE TRANSMIS EN PARAMETRE PAR LE SCRIPT PYTHON.
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
    
    Noeud courant : dans le document comments.xml
    -->
    <xsl:template match="w:comments" mode="comments">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*" mode="comments"/>
            <xsl:apply-templates select="$root/descendant::w:commentReference"/>

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
    
    Noeud courant : dans le document document.xml
    -->
    <xsl:template match="w:commentReference">
        <xsl:choose>
            <!--
            S'il y a déjà un commentaire correspondant à commentReference dans comments.xml :
            on ne fait rien
            -->
            <xsl:when test="contains($commentsId, @w:id)"/>
            
            <!-- Sinon, on crée un commentaire dans comments.xml -->
            <xsl:otherwise>
                <w:comment w:id="{@w:id}" w:author="{$coredocument//dc:creator}"
                    w:date="{$coredocument//dcterms:modified}"
                    w:initials="{substring($coredocument//dc:creator, 1, 1)}">
                    <w:p w:rsidR="{ancestor::w:r[1]/@w:rsidR}"
                        w:rsidRDefault="{ancestor::w:r[1]/@w:rsidR}">
                        <w:pPr>
                            <w:pStyle
                                w:val="{$stylesdocument//w:style[w:name/@w:val='annotation text']/@w:styleId}"
                            />
                        </w:pPr>
                        <w:r>
                            <w:rPr>
                                <w:rStyle
                                    w:val="{$stylesdocument//w:style[w:name/@w:val='annotation reference']/@w:styleId}"
                                />
                            </w:rPr>
                            <w:annotationRef/>
                        </w:r>
                        <w:r>
                            <xsl:attribute name="w:rsidR">
                                <xsl:value-of select="ancestor::w:r[1]/@w:rsidR"/>
                            </xsl:attribute>
                            <w:t>
                                
                                <xsl:variable name="starttext" select="ancestor::w:p/descendant::w:r[1]//w:t[1]/text()"/>
                                
                                <xsl:variable name="endtext" select="ancestor::w:p/descendant::w:r[descendant::w:t][last()]//w:t[last()]/text()"/>
                                
                                <!-- Est-ce que le parahraphe commence par un de ces signes ? -->
                                <xsl:variable name="startCap">
                                    <xsl:choose>
                                        <xsl:when
                                            test="starts-with($starttext, 'A')
                                            or starts-with($starttext, 'B')
                                            or starts-with($starttext, 'C')
                                            or starts-with($starttext, 'D')
                                            or starts-with($starttext, 'E')
                                            or starts-with($starttext, 'F')
                                            or starts-with($starttext, 'G')
                                            or starts-with($starttext, 'H')
                                            or starts-with($starttext, 'I')
                                            or starts-with($starttext, 'J')
                                            or starts-with($starttext, 'K')
                                            or starts-with($starttext, 'L')
                                            or starts-with($starttext, 'M')
                                            or starts-with($starttext, 'N')
                                            or starts-with($starttext, 'O')
                                            or starts-with($starttext, 'P')
                                            or starts-with($starttext, 'Q')
                                            or starts-with($starttext, 'R')
                                            or starts-with($starttext, 'S')
                                            or starts-with($starttext, 'T')
                                            or starts-with($starttext, 'U')
                                            or starts-with($starttext, 'V')
                                            or starts-with($starttext, 'W')
                                            or starts-with($starttext, 'X')
                                            or starts-with($starttext, 'Y')
                                            or starts-with($starttext, 'Z')
                                            or starts-with($starttext, 'À')
                                            or starts-with($starttext, 'Á')
                                            or starts-with($starttext, 'Â')
                                            or starts-with($starttext, 'Ã')
                                            or starts-with($starttext, 'Ä')
                                            or starts-with($starttext, 'Æ')
                                            or starts-with($starttext, 'Ç')
                                            or starts-with($starttext, 'È')
                                            or starts-with($starttext, 'É')
                                            or starts-with($starttext, 'Ê')
                                            or starts-with($starttext, 'Ë')
                                            or starts-with($starttext, 'Ì')
                                            or starts-with($starttext, 'Í')
                                            or starts-with($starttext, 'Î')
                                            or starts-with($starttext, 'Ï')
                                            or starts-with($starttext, 'Ð')
                                            or starts-with($starttext, 'Ñ')
                                            or starts-with($starttext, 'Ò')
                                            or starts-with($starttext, 'Ó')
                                            or starts-with($starttext, 'Ø')
                                            or starts-with($starttext, 'Ù')
                                            or starts-with($starttext, 'Ú')
                                            or starts-with($starttext, 'Û')
                                            or starts-with($starttext, 'Ü')
                                            or starts-with($starttext, 'Ý')
                                            or starts-with($starttext, 'Œ')
                                            or starts-with($starttext, '1')
                                            or starts-with($starttext, '2')
                                            or starts-with($starttext, '3')
                                            or starts-with($starttext, '4')
                                            or starts-with($starttext, '5')
                                            or starts-with($starttext, '6')
                                            or starts-with($starttext, '7')
                                            or starts-with($starttext, '8')
                                            or starts-with($starttext, '9')
                                            or starts-with($starttext, '0')
                                            or starts-with($starttext, '¿')
                                            or starts-with($starttext, '¡')
                                            or starts-with($starttext, '—')
                                            or starts-with($starttext, '-')
                                            or starts-with($starttext, '–')
                                            or starts-with($starttext, '«')">
                                            <xsl:value-of select="'true'"/>
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:value-of select="'false'"/>
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:variable>
                                
                                <!-- Est-ce que le paragraphe finit par un de ces signes ? -->
                                <xsl:variable name="endPunctuation">
                                    <xsl:variable name="endsign"
                                        select="substring($endtext, string-length($endtext))"/>
                                    
                                    <xsl:choose>
                                        <xsl:when
                                            test="$endsign = '.'
                                            or $endsign = '?'
                                            or $endsign = '!'
                                            or $endsign = ':'
                                            or $endsign = ';'
                                            or $endsign = ','
                                            or $endsign = '»'
                                            or $endsign = '…'">
                                            <xsl:value-of select="'true'"/>
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:value-of select="'false'"/>
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:variable>
                                
                                <xsl:choose>
                                    
                                    <!-- Ni majuscule, ni ponctuation finale -->
                                    <xsl:when test="$startCap = 'false' and $endPunctuation = 'false'">
                                        <xsl:text>Paragraphe sans majuscule ni ponctuation finale.</xsl:text>
                                    </xsl:when>
                                    <!-- Une majuscule, mais pas de ponctuation finale -->
                                    <xsl:when test="$startCap = 'true' and $endPunctuation = 'false'">
                                        <xsl:text>Paragraphe sans ponctuation finale.</xsl:text>
                                    </xsl:when>
                                    
                                    <!-- Pas de majuscule, mais ponctuation finale -->
                                    <xsl:when test="$startCap = 'false' and $endPunctuation = 'true'">
                                        <xsl:text>Paragraphe sans majuscule.</xsl:text>
                                    </xsl:when>

                                </xsl:choose>
                            </w:t>
                        </w:r>
                    </w:p>
                </w:comment>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <!-- Pas d'interférence du document source -->
    <xsl:template match="node() | text()"/>
</xsl:stylesheet>
