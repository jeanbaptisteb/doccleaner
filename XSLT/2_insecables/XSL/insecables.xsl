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
    
    
    
    <!-- Copie par défaut. -->
    <xsl:template match="node()[not(self::text())][not(self::w:r)] | @*">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*"/>
        </xsl:copy>
    </xsl:template>
    
    
    
    <!-- non breaking space -->
    <xsl:variable name="nbs">&#160;</xsl:variable>



    <!--
    Cette template traite les segments de texte w:r (text run).
        
    Le texte d'un paragraphe est découpé arbitrairement par MS Word en segments de texte.
    Les césures peuvent se faire n'importe où, par exemple au milieu d'un mot ou juste avant un
    signe qui prendrait une espace insécable.
    
    Ce serait compliqué de gérer les espaces insécables qui se trouveraient à l'endroit  précis
    où le texte est divisé entre deux segments.
    
    Lorsqu'un segment de texte se termine par une espace sécable ou insécable, et que le segment
    de texte suivant commence par une espace sécable ou insécable, ou par un signe qui doit être
    précédé d'une espace insécable : le segment de texte courant est copié sans son espace finale.
    
    L'espace finale supprimée du segment courant sera rajoutée éventuellement au début du segment de
    texte suivant.
    
    Tous les textes des segments, avec ou sans espace finale, sont ensuite copiés en passant par la
    template "manageSpaces" qui gère l'ajout des espaces insécables au sein de la chaîne de caractère.
    -->
    <xsl:template match="w:r">
        <xsl:variable name="nextTextSegment"
            select="following::w:r[ancestor::w:p = current()/ancestor::w:p]/w:t/text()"/>
        <!--            
        1. Si l'élément le segment de texte courant n'est pas le dernier du paragraphe,
           et s'il se termine par une espace sécable ou insécable :
           1.1. Si le segment de texte suivant commence par un signe qui doit être précédé
                par une espace insécable,
                ou s'il commence par une espace sécable ou insécable :
                1.1.1. Si le segment de texte courant ne contient rien qu'une espace sécable ou
                       insécable,
                       et qu'il n'est pas un hyperlien :
                       => On supprime l'élément w:r courant.
                          (L'espace insécable sera rajoutée au début du segment de texte suivant.)
                1.1.2. Si le segment de texte courant ne contient rien qu'une espace sécable ou
                       insécable,
                       et qu'il est un hyperlien (très impprobable) :
                       => On copie le segment de texte courant.
                          (Une espace insécable sera tout de même rajoutée au début du segment de texte suivant.)
                1.1.3. Sinon :
                       => On copie le segment de texte courant en supprimant l'espace à la fin.
                          (L'espace insécable sera rajoutée au début du segment de texte suivant.)
           1.2. Sinon :
                => On copie le segment de texte courant.
        2. Sinon :
           => On copie le segment de texte courant.
           
        La copie du segment de texte courant se fait via la template "manageSpaces" qui rajoute les
        espaces insécables nécessaires, ou qui remplace les espaces sécables en insécables.
        -->
        <xsl:choose>

            <!-- 1. -->
            <xsl:when
                test="following::w:r[ancestor::w:p = current()/ancestor::w:p]
                    and (
                        substring(descendant::w:t//text(), string-length(descendant::w:t//text())) = ' '
                        or substring(descendant::w:t//text(), string-length(descendant::w:t//text())) = $nbs
                    )
                    ">
                <xsl:choose>
                    <!-- 1.1 -->
                    <xsl:when
                        test="starts-with($nextTextSegment, ' ')
                            or starts-with($nextTextSegment, $nbs)
                            or starts-with($nextTextSegment, ':')
                            or starts-with($nextTextSegment, '?')
                            or starts-with($nextTextSegment, ';')
                            or starts-with($nextTextSegment, '!')
                            or starts-with($nextTextSegment, '»')
                            or starts-with($nextTextSegment, '=')
                            or starts-with($nextTextSegment, '%')
                            or starts-with($nextTextSegment, 'p.')
                            or starts-with($nextTextSegment, 'pp.')
                            or starts-with($nextTextSegment, '$')
                            or starts-with($nextTextSegment, '€')
                        ">
                        <xsl:choose>

                            <!-- 1.1.1. -->
                            <xsl:when test="
                                (descendant::w:t//text() = ' ' or descendant::w:t//text() = $nbs)
                                and not(ancestor::w:hyperlink)"/>
                            
                            <!-- 1.1.2. -->
                            <xsl:when test="
                                (descendant::w:t//text() = ' ' or descendant::w:t//text() = $nbs)
                                and ancestor::w:hyperlink">
                                <xsl:copy>
                                    <xsl:apply-templates select="node() | @*"/>
                                </xsl:copy>
                            </xsl:when>

                            <!-- 1.1.3. -->
                            <xsl:otherwise>
                                <xsl:copy>
                                    <xsl:copy-of select="@*"/>
                                    <xsl:apply-templates mode="removeLastSpace"/>
                                </xsl:copy>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>

                    <!-- 1.2. -->
                    <xsl:otherwise>
                        <xsl:copy>
                            <xsl:apply-templates select="node() | @*"/>
                        </xsl:copy>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>

            <!-- 2. -->
            <xsl:otherwise>
                <xsl:copy>
                    <xsl:apply-templates select="node() | @*"/>
                </xsl:copy>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <!--
    Copie classique mais en mode "removeLastSpace",
    pour aboutir à la template de copie du texte sans la dernière espace.
    -->
    <xsl:template match="node()[not(self::text())]" mode="removeLastSpace">
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:apply-templates mode="removeLastSpace"/>
        </xsl:copy>
    </xsl:template>



    <!-- Copie du texte sans la dernière espace. Le texte est envoyé à manageSpaces -->
    <xsl:template match="text()" mode="removeLastSpace">
        <xsl:call-template name="manageSpaces">
            <xsl:with-param name="text" select="substring(., 1, string-length(.) -1)"/>
        </xsl:call-template>
    </xsl:template>



    <!--
    Copie du texte. Le texte est envoyé à manageSpaces.
    Sauf pour les hyperliens auxquels on ne touche pas.
    -->
    <xsl:template match="text()[not(ancestor::w:hyperlink)]">
        <xsl:call-template name="manageSpaces">
            <xsl:with-param name="text" select="."/>
        </xsl:call-template>
    </xsl:template>


    
    <!-- 
    Si le texte n'a pas besoin d'espaces insécables, ou si c'est un lien hypertexte,
    il est simplement copié.
    Sinon, la template ajoute les espaces insécables en 2 temps, avant de copier le texte.
    
    1. Toutes les espaces, sécables ou insécables, qui se trouvent autour de signes qui doivent
       avoir des espaces insécables, sont supprimées.
    2. Les espaces insécables sont rajoutées autour des signes qui doivent en avoir, et qui soit
       n'en avaient pas, soit n'en ont plus depuis 1.
    -->
    <xsl:template name="manageSpaces">
        <xsl:param name="text" select="."/>
        
        <xsl:choose>
            <xsl:when
                test="not(
                contains($text, ':')
                or contains($text, '?')
                or contains($text, ';')
                or contains($text, '!')
                or contains($text, '»')
                or contains($text, '=')
                or contains($text, '%')
                or contains($text, '«')
                or contains($text, 'n°')
                or contains($text, 'op. cit.')
                or contains($text, 'art. cit.')
                or contains($text, ' p.')
                or contains($text, ' pp.')
                or contains($text, '$')
                or contains($text, '€')
                )
                or (starts-with($text, 'http://')
                    or starts-with($text, 'https://')
                    or starts-with($text, 'ftp://')
                    or starts-with($text, 'sftp://')
                    or starts-with($text, 'www.')
                   )
                ">
                <xsl:value-of select="$text"/>
            </xsl:when>

            <xsl:otherwise>
                <xsl:variable name="textWithoutSomeSpaces">
                    <xsl:call-template name="removeSomeSpaces">
                        <xsl:with-param name="text" select="$text"/>
                    </xsl:call-template>
                </xsl:variable>

                <xsl:variable name="finalText">
                    <xsl:call-template name="addNonBreakinSpaces">
                        <xsl:with-param name="text" select="$textWithoutSomeSpaces"/>
                    </xsl:call-template>
                </xsl:variable>

                <xsl:value-of select="$finalText"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>



    <!-- 
    Cette template supprime les espaces sécables ou insécables autour des signes qui doivent avoir
    une espace insécable.
    -->
    <xsl:template name="removeSomeSpaces">
        <xsl:param name="text"/>

        <xsl:variable name="pass1">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$text"/>
                <xsl:with-param name="search" select="' :'"/>
                <xsl:with-param name="replacement" select="':'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass2">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass1"/>
                <xsl:with-param name="search" select="concat($nbs, ':')"/>
                <xsl:with-param name="replacement" select="':'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass3">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass2"/>
                <xsl:with-param name="search" select="' ?'"/>
                <xsl:with-param name="replacement" select="'?'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass4">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass3"/>
                <xsl:with-param name="search" select="concat($nbs, '?')"/>
                <xsl:with-param name="replacement" select="'?'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass5">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass4"/>
                <xsl:with-param name="search" select="' ;'"/>
                <xsl:with-param name="replacement" select="';'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass6">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass5"/>
                <xsl:with-param name="search" select="concat($nbs, ';')"/>
                <xsl:with-param name="replacement" select="';'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass7">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass6"/>
                <xsl:with-param name="search" select="' !'"/>
                <xsl:with-param name="replacement" select="'!'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass8">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass7"/>
                <xsl:with-param name="search" select="concat($nbs, '!')"/>
                <xsl:with-param name="replacement" select="'!'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass9">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass8"/>
                <xsl:with-param name="search" select="' »'"/>
                <xsl:with-param name="replacement" select="'»'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass10">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass9"/>
                <xsl:with-param name="search" select="concat($nbs, '»')"/>
                <xsl:with-param name="replacement" select="'»'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass11">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass10"/>
                <xsl:with-param name="search" select="' ='"/>
                <xsl:with-param name="replacement" select="'='"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass12">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass11"/>
                <xsl:with-param name="search" select="concat($nbs, '=')"/>
                <xsl:with-param name="replacement" select="'='"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass13">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass12"/>
                <xsl:with-param name="search" select="' %'"/>
                <xsl:with-param name="replacement" select="'%'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass14">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass13"/>
                <xsl:with-param name="search" select="concat($nbs, '%')"/>
                <xsl:with-param name="replacement" select="'%'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass15">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass14"/>
                <xsl:with-param name="search" select="'« '"/>
                <xsl:with-param name="replacement" select="'«'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass16">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass15"/>
                <xsl:with-param name="search" select="concat('«', $nbs)"/>
                <xsl:with-param name="replacement" select="'«'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass17">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass16"/>
                <xsl:with-param name="search" select="'n° '"/>
                <xsl:with-param name="replacement" select="'n°'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass18">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass17"/>
                <xsl:with-param name="search" select="concat('n°', $nbs)"/>
                <xsl:with-param name="replacement" select="'n°'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass19">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass18"/>
                <xsl:with-param name="search" select="' p. '"/>
                <xsl:with-param name="replacement" select="' p.'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass20">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass19"/>
                <xsl:with-param name="search" select="concat(' p.', $nbs)"/>
                <xsl:with-param name="replacement" select="' p.'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass21">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass20"/>
                <xsl:with-param name="search" select="' pp. '"/>
                <xsl:with-param name="replacement" select="' pp.'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass22">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass21"/>
                <xsl:with-param name="search" select="concat(' pp.', $nbs)"/>
                <xsl:with-param name="replacement" select="' pp.'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass23">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass22"/>
                <xsl:with-param name="search" select="' $'"/>
                <xsl:with-param name="replacement" select="'$'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass24">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass23"/>
                <xsl:with-param name="search" select="concat($nbs, '$')"/>
                <xsl:with-param name="replacement" select="'$'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass25">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass24"/>
                <xsl:with-param name="search" select="' €'"/>
                <xsl:with-param name="replacement" select="'€'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass26">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass25"/>
                <xsl:with-param name="search" select="concat($nbs, '€')"/>
                <xsl:with-param name="replacement" select="'€'"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="$pass26"/>
    </xsl:template>



    <!-- 
    Cette template ajoute des espaces insécables autour des signes qui doivent en avoir et qui n'en
    avaient pas, ou qui n'en ont plus depuis que le texte est passé par la template
    "removeSomeSpaces"
    -->
    <xsl:template name="addNonBreakinSpaces">
        <xsl:param name="text"/>
        
        <xsl:variable name="pass1">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$text"/>
                <xsl:with-param name="search" select="':'"/>
                <xsl:with-param name="replacement" select="concat($nbs, ':')"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass2">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass1"/>
                <xsl:with-param name="search" select="'?'"/>
                <xsl:with-param name="replacement" select="concat($nbs, '?')"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass3">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass2"/>
                <xsl:with-param name="search" select="';'"/>
                <xsl:with-param name="replacement" select="concat($nbs, ';')"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass4">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass3"/>
                <xsl:with-param name="search" select="'!'"/>
                <xsl:with-param name="replacement" select="concat($nbs, '!')"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass5">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass4"/>
                <xsl:with-param name="search" select="'»'"/>
                <xsl:with-param name="replacement" select="concat($nbs, '»')"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass6">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass5"/>
                <xsl:with-param name="search" select="'='"/>
                <xsl:with-param name="replacement" select="concat($nbs, '=')"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass7">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass6"/>
                <xsl:with-param name="search" select="'%'"/>
                <xsl:with-param name="replacement" select="concat($nbs, '%')"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass8">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass7"/>
                <xsl:with-param name="search" select="'«'"/>
                <xsl:with-param name="replacement" select="concat('«', $nbs)"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass9">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass8"/>
                <xsl:with-param name="search" select="'n°'"/>
                <xsl:with-param name="replacement" select="concat('n°', $nbs)"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass10">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass9"/>
                <xsl:with-param name="search" select="'op. cit.'"/>
                <xsl:with-param name="replacement" select="concat('op.', $nbs, 'cit.')"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass11">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass10"/>
                <xsl:with-param name="search" select="'art. cit.'"/>
                <xsl:with-param name="replacement" select="concat('art.', $nbs, 'cit.')"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass12">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass11"/>
                <xsl:with-param name="search" select="' p.'"/>
                <xsl:with-param name="replacement" select="concat($nbs, 'p.', $nbs)"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass13">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass12"/>
                <xsl:with-param name="search" select="' pp.'"/>
                <xsl:with-param name="replacement" select="concat($nbs, 'pp.', $nbs)"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass14">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass13"/>
                <xsl:with-param name="search" select="'$'"/>
                <xsl:with-param name="replacement" select="concat($nbs, '$')"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="pass15">
            <xsl:call-template name="string-replace">
                <xsl:with-param name="subject" select="$pass14"/>
                <xsl:with-param name="search" select="'€'"/>
                <xsl:with-param name="replacement" select="concat($nbs, '€')"/>
                <xsl:with-param name="global" select="true()"/>
            </xsl:call-template>
        </xsl:variable>

        <xsl:value-of select="$pass15"/>
    </xsl:template>



    <!-- 
    En XPath 1.0, il n'y a pas de fonction replace() comme en XPath 2.0.
    La fonction translate() ne fait que des remplacements de caractères à l'unité.
    
    Cette fonction correspond à la fonction replace() pour du XPath 1.0 natif.
    -->
    <xsl:template name="string-replace">
        <xsl:param name="subject" select="''"/>
        <xsl:param name="search" select="''"/>
        <xsl:param name="replacement" select="''"/>
        <xsl:param name="global" select="false()"/>

        <xsl:choose>
            <xsl:when test="contains($subject, $search)">
                <xsl:value-of select="substring-before($subject, $search)"/>
                <xsl:value-of select="$replacement"/>
                <xsl:variable name="rest" select="substring-after($subject, $search)"/>
                <xsl:choose>
                    <xsl:when test="$global">
                        <xsl:call-template name="string-replace">
                            <xsl:with-param name="subject" select="$rest"/>
                            <xsl:with-param name="search" select="$search"/>
                            <xsl:with-param name="replacement" select="$replacement"/>
                            <xsl:with-param name="global" select="$global"/>
                        </xsl:call-template>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="$rest"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$subject"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

</xsl:stylesheet>