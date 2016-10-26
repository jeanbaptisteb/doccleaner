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
    
    
    
    <xsl:param name="tempdir"/>    
    <xsl:variable name="tempdirSlash" select="translate($tempdir, '\', '/')"/>



    <!-- Copie par défaut. -->
    <xsl:template match="node() | @*">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>
        </xsl:copy>
    </xsl:template>
    
    
    
    <xsl:template match="w:p//w:r[last()]">
        <!-- 
        Si on a un segment de texte w:r (text run), descendant en dernière position
        d'un paragraphe de titre ou intertitre w:p.
    
        "Descendant en dernière position" parce que c'est à la fin du dernier segment de texte
        qu'on trouve le point final éventuel à supprimer.
    
        Pour trouver les paragraphes stylés en titre ou intertitres,
        sans utiliser le nom du style en français :
    
        On fait correspondre le nom du style dans document.xml (w:pStyle/@w:val), (Titre, Soustitre, Titre1...),
        au style dans styles.xml (w:style/w:name/@w:val), dont le nom du style générique en anglais 
        correspond à un titre ou à un intertitre (Title, Subtitle, heading 1...)
        -->
        <xsl:choose>
            <xsl:when test="ancestor::w:p[w:pPr[w:pStyle[
                @w:val = document(concat($tempdirSlash, '/word/styles.xml'))//w:style[
                w:name/@w:val = 'Title'
                or w:name/@w:val = 'Surtitre'
                or w:name/@w:val = 'Subtitle'
                or w:name/@w:val = 'heading 1'
                or w:name/@w:val = 'heading 2'
                or w:name/@w:val = 'heading 3'
                or w:name/@w:val = 'heading 4'
                or w:name/@w:val = 'heading 5'
                or w:name/@w:val = 'heading 6'
                or w:name/@w:val = 'heading 7'
                or w:name/@w:val = 'heading 8'
                or w:name/@w:val = 'heading 9'
                ]/@w:styleId
                ]]]">
                <xsl:choose>
                    <!-- Si le texte n'est qu'un point, on supprime l'élément courant w:r -->
                    <xsl:when test="descendant::w:t//text() = '.'"/>
                    
                    <!--
                    Si le texte se termine par un point, on copie tous les éléments entre w:r et le texte.
                    Pour à la fin copier le texte dans une template spéciale qui enlève le point final.
                    -->
                    <xsl:when
                        test="substring(descendant::w:t//text(), string-length(descendant::w:t//text())) = '.'">
                        <xsl:copy>
                            <xsl:copy-of select="@*"/>
                            <xsl:apply-templates mode="removePoint"/>
                        </xsl:copy>
                    </xsl:when>
                    
                    <!-- Sinon on copie tout normalement -->
                    <xsl:otherwise>
                        <xsl:copy>
                            <xsl:apply-templates select="node() | @*"/>
                        </xsl:copy>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            
            <!--
            Si le segment de texte ne fait pas partie d'un titre ou d'un intertitre,
            on le copie normalement
            -->
            <xsl:otherwise>
                <xsl:copy>
                    <xsl:apply-templates select="node() | @*"/>
                </xsl:copy>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    
    
    <!--
    Copie classique mais en mode "removePoint",
    pour aboutir à la template de copie du texte sans le point final.
    -->
    <xsl:template match="node()[not(self::text())] | @*" mode="removePoint">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*" mode="removePoint"/>
        </xsl:copy>
    </xsl:template>



    <!-- Copie du texte sans le point final. -->
    <xsl:template match="text()" mode="removePoint">
        <xsl:value-of select="substring(., 1, string-length(.) -1)"/>
    </xsl:template>
</xsl:stylesheet>
