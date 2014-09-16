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
    
    
    
    <!--
    Transformation XSLT non dynamique. Tous les styles traités sont écrits en dur.
    
    - J'ai testé la génération dynamique d'expressions XPath, qui sont ensuite stockées dans des
    variables. Mais ces variables ne peuvent pas être évaluées par le processeur XSLT (1.0 ou 2.0).
    - Il existe des extensions pour évaluer dynamiquement du XPath mais elles ne sont supportées
    ni par lxml ni par Saxon.
    - Ecrire du XPath dynamique rend cette transformation XSLT totalement illisible, vu qu'on ne
    comprend plus trop ce qui est évalué.
    -->
    
    
    
    <!--
    Copie de tous les noeuds et attributs.
    Sauf les noeuds <w:r> (dont les fils sont <w:t>) qui font partie des paragraphes
    "description auteur" ou "résumé" qui en suivent immédiatement d'autres.
    Ceux-là sont traités dans des templates particulières plus bas.
    -->
    <xsl:template
        match="node()[not(w:t[ancestor::w:p[w:pPr/w:pStyle/@w:val='DescriptionAuteur']
        [preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'DescriptionAuteur']])]
        [not(w:t[ancestor::w:p[w:pPr/w:pStyle/@w:val='Resume']
        [preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'Resume']])]
        [not(w:t[ancestor::w:p[w:pPr/w:pStyle/@w:val='Abstract']
        [preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'Abstract']])]
        | @*">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>
        </xsl:copy>
    </xsl:template>
    
    
    
    <!--
    Sélection du dernier <w:r> des paragraphes de description d'auteurs ou de résumés qui ne
    doivent eux-mêmes pas être fusionnés.
    
    A l'intérieur de ce dernier <w:r>, on copie (2nd apply-templates) le texte des paragraphes
    suivants qui doivent éventuellement être fusionés.
    -->
    <xsl:template
        match="w:p[w:pPr/w:pStyle/@w:val='DescriptionAuteur']
        [not(preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'DescriptionAuteur')]//w:r[last()]
        ">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>

            <xsl:apply-templates
                select="ancestor::w:p/following-sibling::w:p[w:pPr/w:pStyle/@w:val = 'DescriptionAuteur']
                [not(preceding-sibling::w:p[not(w:pPr/w:pStyle/@w:val='DescriptionAuteur')]
                [preceding-sibling::*[self::* = current()/ancestor::w:p]])
                ]//w:r/w:t"
            />
        </xsl:copy>
    </xsl:template>
    
    
    
    <xsl:template
        match="w:p[w:pPr/w:pStyle/@w:val='Resume']
        [not(preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'Resume')]//w:r[last()]
        ">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>

            <xsl:apply-templates
                select="ancestor::w:p/following-sibling::w:p[w:pPr/w:pStyle/@w:val = 'Resume']
                [not(preceding-sibling::w:p[not(w:pPr/w:pStyle/@w:val='Resume')]
                [preceding-sibling::*[self::* = current()/ancestor::w:p]])
                ]//w:r/w:t"
            />
        </xsl:copy>
    </xsl:template>
    
    
    
    <xsl:template
        match="w:p[w:pPr/w:pStyle/@w:val='Abstract']
        [not(preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'Abstract')]//w:r[last()]
        ">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>

            <xsl:apply-templates
                select="ancestor::w:p/following-sibling::w:p[w:pPr/w:pStyle/@w:val = 'Abstract']
                [not(preceding-sibling::w:p[not(w:pPr/w:pStyle/@w:val='Abstract')]
                [preceding-sibling::*[self::* = current()/ancestor::w:p]])
                ]//w:r/w:t"
            />
        </xsl:copy>
    </xsl:template>
    
    
    
    <!--
    Le texte des paragraphes qui doivent être fusionnés.
    On le copie avec une balise <w:br> avant pour qu'il y ait un saut de ligne.
    -->
    <xsl:template
        match="w:t[ancestor::w:p[w:pPr/w:pStyle/@w:val='DescriptionAuteur']
        [preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'DescriptionAuteur']]
        | w:t[ancestor::w:p[w:pPr/w:pStyle/@w:val='Resume']
        [preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'Resume']]
        | w:t[ancestor::w:p[w:pPr/w:pStyle/@w:val='Abstract']
        [preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'Abstract']]
        ">
        <w:br/>
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>
        </xsl:copy>
    </xsl:template>
    
    
    
    <!--
    On ne sélectionne pas les paragraphes dont le texte a été fusionné avec le paragraphe précédent.
    -->
    <xsl:template
        match="w:p[w:pPr/w:pStyle/@w:val='DescriptionAuteur']
        [preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'DescriptionAuteur']
        | w:p[w:pPr/w:pStyle/@w:val='Resume']
        [preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'Resume']
        | w:p[w:pPr/w:pStyle/@w:val='Abstract']
        [preceding-sibling::w:p[1]/w:pPr/w:pStyle/@w:val = 'Abstract']
        "
    />
</xsl:stylesheet>
