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
    
    
    
    <!-- On copie tout sauf les paragraphes vides (sauts de paragraphes vides). -->
    <xsl:template match="node()| @*">
        <xsl:choose>
            <xsl:when test="self::w:p
                and (parent::w:body
                or parent::w:footnote
                or parent::w:endnote)
                and not(descendant::w:t)
                and not(descendant::w:drawing)"/>
            <xsl:otherwise>
                <xsl:copy>
                    <xsl:apply-templates select="node() | @*"/>
                </xsl:copy>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

</xsl:stylesheet>