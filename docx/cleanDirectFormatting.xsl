<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
    xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">

<!-- Cleaning up output XML -->
<xsl:strip-space elements="*" />
<xsl:output method="xml" indent="yes" />

<!-- Copying original content -->  
<xsl:template match="node() | @*">
    <xsl:copy>
        <xsl:apply-templates select="node() | @*"/>
    </xsl:copy>
</xsl:template>

<!-- List of tags we keep: -->
    <xsl:template match="*/w:rPr">
		<xsl:copy>
			<xsl:copy-of select="w:b" /> <!-- bold -->
			<xsl:copy-of select="w:i" /> <!-- italics -->
			<xsl:copy-of select="w:u" /> <!-- underline -->
			<xsl:copy-of select="w:vertAlign" /> <!-- superscript and subscript -->
			<xsl:copy-of select="w:lang" /> <!-- language -->
		</xsl:copy>
	</xsl:template>
</xsl:stylesheet>
