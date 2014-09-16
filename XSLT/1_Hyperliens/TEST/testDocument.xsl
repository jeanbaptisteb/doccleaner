<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">



	<xsl:template match="processedDocument">
		<xsl:copy>
			<xsl:apply-templates select="node() | @*"/>

			<!--
			<xsl:apply-templates
				select="document($path)//para"/>
			-->
			
			<xsl:apply-templates
				select="document('../ReadThis.xml', /)//para"/>
		</xsl:copy>
	</xsl:template>



	<xsl:template match="p">
		<pa>
			<xsl:apply-templates/>
		</pa>
	</xsl:template>



	<xsl:template match="para">
		<pa>
			<xsl:apply-templates/>
		</pa>
	</xsl:template>
</xsl:stylesheet>
