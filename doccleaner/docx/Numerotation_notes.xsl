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



    <xsl:variable name="endOrFootnotes" select="local-name(/node())"/>
    <xsl:variable name="note">
        <xsl:choose>
            <xsl:when test="$endOrFootnotes = 'endnotes'">
                <xsl:value-of select="'endnote'"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="'footnote'"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>

    <xsl:variable name="numbers" select="'1234567890_-. '"/>



    <!-- 
	Le script Python passe dans $tempdir le chemin absolu vers le répertoire temporaire où se trouve le fichier
	XML source.
	La lib Python lxml ne gère pas bien la fonction XPath document() et n'accepte que des chemins
	relatifs à la feuille XSLT, ou des chemins absolus (2ème argument de la fonction document() non géré).
	-->
    <xsl:param name="tempdir"/>
    <xsl:variable name="tempdirSlash" select="translate($tempdir, '\', '/')"/>
    
    
    
    
    <xsl:variable name="stylesdocument" select="document(concat($tempdirSlash,'/word/styles.xml'))"/>
    <!--
    <xsl:variable name="stylesdocument" select="document('styles.xml', /)"/>
    -->
    
    
    
    <!--
    Cette transformation XSLT répare les notes cassées. Il faut vraiment que la note soit dans un
    sale état pour qu'elle ne soit pas bien réparée.
    En particulier, elle nettoie les numéros des appels de note au début des notes :
    - Certains numéros de notes sont inscrits en dur, plutôt que d'être incrémentés automatiquement.
    - Certains numéros de notes sont absents.
      ==> La template rétablit les numéros de note automatiques, et supprime les numéros de notes
          manuels le cas échéant.
    
    Une note est encodée en 3 parties comme suit :
    ==============================================================================================
    <w:footnote>
        <w:p> 
    =====  1ère partie : style de la note, immuable  =============================================
            <w:pPr>
                <w:pStyle w:val="Notedebasdepage"/>
            </w:pPr>
    ===== 2ème partie : style du numéro de note, et numéro de note automatique, immuable =========
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="Appelnotedebasdep"/>
                </w:rPr>
                <w:footnoteRef/>
            </w:r>
    ===== 3ème partie : textes de la note =========================================================
            <w:r>
                <w:t>Texte de la note</w:t>
                <w:t>Suite du texte de la note</w:t>
            </w:r>
            <w:r>
                <w:t>Suite du texte de la note.</w:t>
            </w:r> 
        </w:p>
    </w:footnote>
    ==============================================================================================
    
    I. Pour chaque w:p, on va vérifier les 3 parties de la note :
    
    1. On copie w:pPr s'il existe, sinon on le crée.
    
    2. - Si le style du numéro de note est bien formé (s'il existe et qu'il contient une balise de
         numéro de note automatique) : on le copie.
       - S'il n'est pas bien formé : on le construit avec une balise d'appel de note automatique.
         (w:r[w:rPr], avec pour fils  w:r/w:rPr  et w:r/wfootnoteRef)
           
    3. Ensuite, on copie le texte de la note. Cf. template suivante.
       (w:r["qui n'ont pas de descendant '@w:val=Appelnotedebasdep'"])
    
    
    II. Pour chaque w:r :
        Pour les w:r de la troisième partie, ceux qui contiennent le texte de la note, on ne veut
        pas que ce texte commence par un numéro de note manuel.
        Il faut tester toutes les fantaisies d'encodage de Word. Le numéro de note manuel peut être
        encodé :
        a. Isolément dans une balise w:t
        <w:r>
           <w:t>1. </w:t>
        </w:r>
        <w:r>
          <w:t>Texte de la note</w:t>
        </w:r>
        
        b. Dans plusieurs balises w:t descendants de la même balise w:r
        <w:r>
          <w:t>1</w:t>
          <w:t>5</w:t>
        <w:r>
        <w:r>
          <w:t> Texte de la 15ème note.<w:t>
        <w:r>
        
        c. Dans plusieurs balises w:r
        <w:r>
          <w:t>1</w:t>
        </w:r>
        <w:r>
          <w:t>2</w:t>
        </w:r>
          <w:t> Texte de la douxième note.</w:t>
        </w:r>
        
        d. Au début d'une chaîne de caractère légitime
        <w:r>
          <w:t>10. La dixième note</w:t>
        </w:r>
        
        e. Un mélange de tout.
        <w:r>
          <w:t>1</w:t>
          <w:t>2</w:t>
        </w:r>
        <w:r>
          <w:t>3</w:t>
          <w:t>4. La 1234ème note</w:t>
        </w:t>
        
        f. Ne pas supprimer (pas d'espace après le numéro)
        <w:r>
           <w:t> 3ème place ex-aequo.</w:t>
        </w:r>
        
        g. Correct également (un w:t précédent le chiffre n'a pas été supprimé). :
        <w:r>
           <w:t> L'émission </w:t>
           <w:t>30 </w:t>
           <w:t>millions d'amis.</w:t>
        </w:r>
        
        On teste les w:r.
        1. Si, au sein du même ancêtre w:p, une balise w:t d'un w:r précédent n'a
           pas été supprimée, c'est qu'elle contenait du texte légitime. Donc tout le texte qui
           vient après est ok. Donc on copie la balise w:r courante.
        2. Si, au sein du même ancêtre w:p, il n'y a pas eu de balise w:t précédente qui a été
           gardée, le texte de la balise w:r courante commence peut-être par un numéro de note.
           2.1. Si le texte de la balise w:r courante ne contient que des chiffres, c'est un appel
                de note manuel. On ne copie pas la balise w:r.
           2.2. Sinon, on copie la balise w:r. On testera les descendants w:t dans une template
                séparée.
    
    
    III. Pour chaque w:t
         1. Si, au sein du même ancêtre w:p, une balise w:t précédente n'a pas été supprimée, c'est
            qu'elle contenait du texte légitime. Donc tout le texte qui vient après est ok. Donc on
            copie la balise w:t courante.
         2. Si, au sein du même ancêtre w:p, il n'y a pas eu de balise w:t précédente qui a été
            gardée, le texte de la balise w:t courante est peut-être un numéro de note manuel.
            2.1. Si le texte de la balise w:t courante ne contient que des chiffres, c'est un appel
                 de note manuel. On ne copie pas la balise w:t courante.
            2.2. Sinon, on copie la balise w:t. On testera la chaîne de caractère dans une template
                 spécifique pour vérifier qu'elle ne commence pas par un numéro.
                 
                 
    IV. Pour chaque text() descendant d'une balise w:t
        1. Si, au sein du même ancêtre w:p, une balise w:t précédente n'a pas été supprimée, c'est
           qu'elle contenait du texte légitime. Donc tout le texte qui vient après est ok. On copie
           le texte courant.
        2. Sinon, on copie le texte après les chiffres qu'il y a éventuellement au début de la chaîne
           de caractère.
    -->



    <!-- Copie par défaut. -->
    <xsl:template match="node() | @*">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>
        </xsl:copy>
    </xsl:template>



    <xsl:template
        match="w:p[not(descendant::w:separator)][not(descendant::w:continuationSeparator)]">
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            
            <!-- Première partie d'une note : style de la note -->
            <xsl:choose>
                <xsl:when test="w:pPr">
                    <xsl:apply-templates select="w:pPr"/>
                </xsl:when>
                <xsl:otherwise>
                    <w:pPr>
                        <w:pStyle
                            w:val="{$stylesdocument//w:style[w:name/@w:val =
                            concat($note, ' text')]/@w:styleId}"
                        />
                    </w:pPr>
                </xsl:otherwise>
            </xsl:choose>
            
            <!--
            Deuxième partie d'une note : numéro de la note et style du numéro de la note

            On test si cette partie est bien formée : s'il y a un style "appel de note", 
            et un numéro de note automatique.
            -->
            <xsl:variable name="note_number">
                <xsl:choose>
                    <xsl:when test="$note = 'footnote'">
                        <xsl:choose>
                            <xsl:when
                                test="(w:r/w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                            'footnote reference']/@w:styleId)
                            and w:r[w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                            'footnote reference']/@w:styleId]/descendant::w:footnoteRef">
                                <xsl:value-of select="'true'"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="'false'"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:choose>
                            <xsl:when
                                test="(descendant::w:r/w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                                'endnote reference']/@w:styleId)
                                and descendant::w:r[w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                                'endnote reference']/@w:styleId]/descendant::w:endnoteRef">
                                <xsl:value-of select="'true'"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="'false'"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:variable>

            <xsl:choose>
                <!-- Si le w:r de l'appel de note est bien formé, le on copie -->
                <xsl:when test="$note_number = 'true'">
                    <xsl:choose>
                        <xsl:when test="$note = 'footnote'">
                            <xsl:apply-templates
                                select="descendant::w:r[w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                                'footnote reference']/@w:styleId]"
                            />
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:apply-templates
                                select="descendant::w:r[w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                                'endnote reference']/@w:styleId]"
                            />
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:when>

                <!-- Sinon, on le recrée avec le numéro de note automatique -->
                <xsl:otherwise>
                    <w:r>
                        <xsl:choose>
                            <xsl:when test="$note = 'footnote'">
                                <w:rPr>
                                    <w:rStyle
                                        w:val="{$stylesdocument//w:style[w:name/@w:val = 'footnote reference']/@w:styleId}"
                                    />
                                </w:rPr>
                                <w:footnoteRef/>
                            </xsl:when>
                            <xsl:otherwise>
                                <w:rPr>
                                    <w:rStyle
                                        w:val="{$stylesdocument//w:style[w:name/@w:val = 'endnote reference']/@w:styleId}"
                                    />
                                </w:rPr>
                                <w:endnoteRef/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </w:r>
                </xsl:otherwise>
            </xsl:choose>

            <!-- 3ème partie de la note : texte. Appel des templates des balises filles -->
            <xsl:choose>
                <xsl:when test="$note = 'footnote'">
                    <xsl:apply-templates
                        select="descendant::w:r[not(w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                        'footnote reference']/@w:styleId)][not(descendant::w:continuationSeparator)][not(descendant::w:separator)][not(descendant::w:footnoteRef)][not(descendant::w:endnoteRef)]"
                    />
                </xsl:when>
                <xsl:otherwise>
                    <xsl:apply-templates
                        select="descendant::w:r[not(w:rPr/w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                        'endnote reference']/@w:styleId)][not(descendant::w:continuationSeparator)][not(descendant::w:separator)][not(descendant::w:footnoteRef)][not(descendant::w:endnoteRef)]"
                    />
                </xsl:otherwise>
            </xsl:choose>
        </xsl:copy>
    </xsl:template>



    <xsl:template match="w:r[not(descendant::w:continuationSeparator)][not(descendant::w:separator)][not(descendant::w:footnoteRef)][not(descendant::w:endnoteRef)]">
        <!--
        On cherche si dans les w:t précédents ayant le même ancêtre w:p que le w:r courant, il
        y a du texte qui ne doit pas être supprimé (qui n'est pas un numéro de note manuel).
        -->
        <xsl:variable name="preceding-suppressed-t-node">
            <xsl:for-each
                select="preceding::w:r[not(descendant::w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                concat($note, ' reference')]/@w:styleId)][ancestor::w:p = current()/ancestor::w:p]/descendant::w:t/descendant::text()
                ">
                <xsl:choose>
                    <xsl:when test="translate(., $numbers, '') = ''">
                        <xsl:value-of select="'true'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="'false'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each>
        </xsl:variable>

        <xsl:choose>
            <!--
            Si parmi les w:t précédents, du même ancêtre w:p, il y a du texte qui ne doit pas être
            supprimé, le texte qui vient après ne constitue pas un numéro de note manuel : on le
            copie.
            -->
            <xsl:when test="contains($preceding-suppressed-t-node, 'false')">
                <xsl:copy>
                    <xsl:apply-templates select="node() | @*"/>
                </xsl:copy>
            </xsl:when>
            
            <!-- Sinon on vérifie si le w:r courant contient du texte qui ne doit pas être supprimé. -->
            <xsl:otherwise>
                <xsl:variable name="descendant-suppressed-t-node">
                    <xsl:for-each select="descendant::w:t/descendant::text()">
                        <xsl:choose>
                            <xsl:when test="translate(., $numbers, '') = ''">
                                <xsl:value-of select="'true'"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="'false'"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:for-each>
                </xsl:variable>

                <xsl:choose>
                    <!-- Si le w:r courant contient du texte qui ne doit pas être supprimé, on le
                        copie -->
                    <xsl:when test="contains($descendant-suppressed-t-node, 'false')">
                        <xsl:copy>
                            <xsl:apply-templates select="node() | @*"/>
                        </xsl:copy>
                    </xsl:when>
                    <!-- Sinon rien, il n'est pas copié. -->
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <xsl:template match="w:t">
        <!--
        On cherche si dans les w:t précédents ayant le même ancêtre w:p que le w:t courant, il
        y a du texte qui ne doit pas être supprimé (qui n'est pas un numéro de note manuel).
        -->
        <xsl:variable name="preceding-suppressed-t-node">
            <xsl:for-each
                select="ancestor::w:r/preceding::w:r[not(descendant::w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                concat($note, ' reference')]/@w:styleId)][ancestor::w:p = current()/ancestor::w:p]/descendant::w:t/descendant::text()
                ">
                <xsl:choose>
                    <xsl:when test="translate(., $numbers, '') = ''">
                        <xsl:value-of select="'true'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="'false'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each>
        </xsl:variable>

        <xsl:choose>
            <!--
            Si parmi les w:t précédents, du même ancêtre w:p, il y a du texte qui ne doit pas être
            supprimé, le texte qui vient après ne constitue pas un numéro de note manuel : on le
            copie.
            -->
            <xsl:when test="contains($preceding-suppressed-t-node, 'false')">
                <xsl:copy>
                    <xsl:apply-templates select="node() | @*"/>
                </xsl:copy>
            </xsl:when>
            
            <!--
            Sinon s'il ne contient pas que des chiffres (et n'est donc pas complètement un numéro de
            note manuel) : on le copie.
            Une template dédiée vérifiera si la chaîne de caractère commence par un numéro de
            note manuel et a besoin d'être traitée.
            -->
            <xsl:otherwise>
                <xsl:choose>
                    <xsl:when test="translate(descendant::text(), $numbers, '') != ''">
                        <xsl:copy>
                            <xsl:apply-templates select="node() | @*"/>
                        </xsl:copy>
                    </xsl:when>
                    <!-- Sinon rien, il n'est pas copié. -->
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <xsl:template match="text()[ancestor::w:t]">
        <!--
        On cherche si dans les w:t précédents ayant le même ancêtre w:p que le texte courant, il
        y a du texte qui ne doit pas être supprimé (qui n'est pas un numéro de note manuel).
        -->
        <xsl:variable name="preceding-suppressed-t-node">
            <xsl:for-each
                select="ancestor::w:r/preceding::w:r[not(descendant::w:rStyle/@w:val = $stylesdocument//w:style[w:name/@w:val =
                concat($note, ' reference')]/@w:styleId)][ancestor::w:p = current()/ancestor::w:p]/descendant::w:t/descendant::text()
                ">
                <xsl:choose>
                    <xsl:when test="translate(., $numbers, '') = ''">
                        <xsl:value-of select="'true'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="'false'"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each>
        </xsl:variable>

        <xsl:choose>
            <!--
            Si parmi les noeuds textes précédents, du même ancêtre w:p, il y a du texte qui ne doit
            pas être supprimé, le texte qui vient après ne constitue pas un numéro de note manuel :
            on le copie.
            -->
            <xsl:when test="contains($preceding-suppressed-t-node, 'false')">
                <xsl:value-of select="."/>
            </xsl:when>
            <!-- Sinon on checke le début de la chaîne de caractère -->
            <xsl:otherwise>
                <xsl:choose>
                    <!--
                    Si le texte commence par des chiffres, suivis d'un espace :
                    ==> On supprime les chiffres (considérés comme la suite de l'appel de note manuel).
                    -->
                    <xsl:when
                        test="contains(normalize-space(.), ' ')
                        and translate(substring-before(normalize-space(.), ' '), $numbers, '') = ''">
                        <xsl:text> </xsl:text>
                        <xsl:value-of select="substring-after(normalize-space(.), ' ')"/>
                    </xsl:when>
                    <!-- Sinon on copie normalement -->
                    <xsl:otherwise>
                        <xsl:value-of select="."/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:stylesheet>
