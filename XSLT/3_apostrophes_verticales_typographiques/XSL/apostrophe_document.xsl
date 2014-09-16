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
	
	
	
	<xsl:template match="node() | @*">
		<xsl:copy>
			<xsl:apply-templates select="node() | @*"/>
		</xsl:copy>
	</xsl:template>
	
	
	
	<xsl:template match="text()">
		<xsl:variable name="apos">'</xsl:variable>
	    <!-- <xsl:variable name="typo_apos">’</xsl:variable> -->
	    <xsl:variable name="typo_apos">&#8217;</xsl:variable>
		
		<xsl:value-of select="translate(., $apos, $typo_apos)"/>
	</xsl:template>
</xsl:stylesheet>