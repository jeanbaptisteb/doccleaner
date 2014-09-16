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


    <!--
    <xsl:param name="open_quotes" select="'“'"/>
    <xsl:param name="closing_quotes" select="'”'"/>
    -->

    <xsl:variable name="quotes">"</xsl:variable>
    <xsl:param name="open_quotes" select="'«'"/>
    <xsl:param name="closing_quotes" select="'»'"/>
    <xsl:variable name="nbs">&#160;</xsl:variable>
    
    <xsl:variable name="open">
        <xsl:choose>
            <xsl:when test="$open_quotes = '«' and not(contains($open_quotes, $nbs))">
                <xsl:value-of select="concat($open_quotes, $nbs)"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$open_quotes"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    <xsl:variable name="close">
        <xsl:choose>
            <xsl:when test="$closing_quotes = '»' and not(contains($closing_quotes, $nbs))">
                <xsl:value-of select="concat($nbs, $closing_quotes)"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$closing_quotes"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>



    <!-- Copie par défaut. -->
    <xsl:template
        match="node() | @*">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>
        </xsl:copy>
    </xsl:template>



    <!--
    Copie des champs de texte.
    
    On compte le nombre de guillemets dans les champs de texte précédents, du même paragraphe.
    (Il y a une template "counter" de comptage.)
    Si c'est un nombre pair, le dernier est fermé, donc le premier guillemet de ce champ devra
    être ouvert.
    Si c'est un nombre impair, le dernier est ouvert, donc le premier guillemet de ce champ devra
    être fermé.
    A priori ces opérations de comptage nécessitent beaucoup de calcul et peuvent ralentir
    l'exécution de la transformation...
    
    Ensuite le texte est copié via une template copyText qui remplace les guillemets droits par des
    guillemets ouvrants et fermants alternativement, en commençant par un guillemet ouvrant ou 
    fermant en fonction de ce qui lui a été transmis en paramètre.
    -->
    <xsl:template match="w:t/text()">
        <xsl:choose>
            <xsl:when test="contains(preceding::text()[parent::w:t][ancestor::node() =
                current()/ancestor::w:p], $quotes) or contains(., $quotes)">
                
                <xsl:variable name="lastIsOpenOrClose">
                    <xsl:variable name="nbPrecedingQuotationMarks">
                        <xsl:call-template name="countQuotationMark">
                            <xsl:with-param name="text">
                                <xsl:for-each select="preceding::text()[parent::w:t][ancestor::node() =
                                    current()/ancestor::w:p]">
                                    <xsl:value-of select="."/>
                                </xsl:for-each>
                                
                            </xsl:with-param>
                        </xsl:call-template>
                    </xsl:variable>
                    
                    <xsl:choose>
                        <xsl:when test="$nbPrecedingQuotationMarks mod 2 = 0">
                            <xsl:value-of select="'close'"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:value-of select="'open'"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:variable>
                
                <xsl:call-template name="copyText">
                    <xsl:with-param name="text" select="."/>
                    <xsl:with-param name="lastIsOpenOrClose" select="$lastIsOpenOrClose"/>
                </xsl:call-template>
                
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="."/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <xsl:template name="copyText">
        <xsl:param name="text"/>
        <xsl:param name="lastIsOpenOrClose"/>
        <xsl:choose>

            <xsl:when test="contains($text, $quotes)">
                <xsl:value-of select="substring-before($text, $quotes)"/>

                <xsl:choose>
                    <xsl:when test="$lastIsOpenOrClose = 'close'">
                        <xsl:value-of select="$open"/>
                    </xsl:when>

                    <xsl:otherwise>
                        <xsl:value-of select="$close"/>
                    </xsl:otherwise>
                </xsl:choose>

                <xsl:call-template name="copyText">
                    <xsl:with-param name="text">
                        <xsl:value-of select="substring-after($text, $quotes)"/>
                    </xsl:with-param>

                    <xsl:with-param name="lastIsOpenOrClose">
                        <xsl:choose>
                            <xsl:when test="$lastIsOpenOrClose = 'close'">
                                <xsl:value-of select="'open'"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="'close'"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:with-param>

                </xsl:call-template>
            </xsl:when>

            <xsl:otherwise>
                <xsl:value-of select="$text"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <xsl:template name="countQuotationMark">
        <xsl:param name="text"/>
        <xsl:param name="counter" select="0"/>
        <xsl:choose>
            <xsl:when test="contains($text, $quotes)">
                <xsl:call-template name="countQuotationMark">
                    <xsl:with-param name="text" select="substring-after($text, $quotes)"/>
                    <xsl:with-param name="counter" select="$counter + 1"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$counter"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:stylesheet>
