function ConvertFrom-MarkdownToHtml {
    <#
        .SYNOPSIS
        Converts Markdown text to HTML with support for common Markdown syntax.

        .DESCRIPTION
        Lightweight Markdown to HTML converter supporting headers, lists, tables, code blocks,
        links, images, bold, italic, blockquotes, and horizontal rules.

        .PARAMETER MarkdownText
        The Markdown text to convert to HTML.

        .EXAMPLE
        PS C:\> ConvertFrom-MarkdownToHtml -MarkdownText "# Hello World`n`nThis is **bold** text."

        .OUTPUTS
        System.String. Returns HTML string.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$MarkdownText
    )

    # Input validation
    if ([string]::IsNullOrEmpty($MarkdownText)) {
        return ""
    }

    $MarkdownText = $MarkdownText.Trim()
    $html = $MarkdownText

    # Normalize line endings to \n only (remove \r)
    $html = $html -replace "`r`n", "`n"
    $html = $html -replace "`r", "`n"

    # Escape Markdown characters first
    $html = $html -replace '\\(.)', '§ESCAPED§$1§ESCAPED§'

    # Extract and protect code blocks before processing other markdown elements
    # This prevents headers and other markdown syntax inside code blocks from being transformed
    # Store code blocks in an array and replace them with placeholders
    $codeBlocks = @()
    $codeBlockIndex = 0

    # Extract code blocks with language support (handles both ``` and malformed ` variants)
    # Note: Some markdown content may have malformed code blocks with single backtick instead of triple backticks
    # (e.g., `powershell instead of ```powershell). This regex handles both cases by matching 1-3 backticks
    # at the start of a line (with optional indentation).
    $html = $html -replace '(?sm)^\s*`{1,3}(\w+)?\r?\n(.+?)^\s*`{1,3}\s*$', {
        $language = $_.Groups[1].Value
        $code = $_.Groups[2].Value -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '\\`', '`'

        $htmlBlock = if ($language) {
            "<pre><code class=`"language-$language`">$code</code></pre>"
        }
        else {
            "<pre><code>$code</code></pre>"
        }

        $placeholder = "§CODEBLOCK§$codeBlockIndex§"
        $codeBlocks += $htmlBlock
        $codeBlockIndex++
        return $placeholder
    }

    # Extract and protect inline code before processing other markdown
    $inlineCodeBlocks = @()
    $inlineCodeIndex = 0
    $html = $html -replace '`([^`]+)`', {
        $code = $_.Groups[1].Value -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '\\`', '`'
        $htmlInline = "<code>$code</code>"

        $placeholder = "§INLINECODE§$inlineCodeIndex§"
        $inlineCodeBlocks += $htmlInline
        $inlineCodeIndex++
        return $placeholder
    }

    # Horizontal rules
    $html = $html -replace '(?m)^(-{3,}|\*{3,}|_{3,})$', '<hr />'

    # Headers (all 6 levels) - now safe from code block interference
    # Also supports headers without space after # (e.g., #Header instead of # Header)
    $html = $html -replace '(?m)^######\s*(.+)$', '<h6>$1</h6>'
    $html = $html -replace '(?m)^#####\s*(.+)$', '<h5>$1</h5>'
    $html = $html -replace '(?m)^####\s*(.+)$', '<h4>$1</h4>'
    $html = $html -replace '(?m)^###\s*(.+)$', '<h3>$1</h3>'
    $html = $html -replace '(?m)^##\s*(.+)$', '<h2>$1</h2>'
    $html = $html -replace '(?m)^#\s*(.+)$', '<h1>$1</h1>'

    # Bold and Italic (limit to single line to prevent backtracking)
    $html = $html -replace '\*\*([^\n\r*]+)\*\*', '<strong>$1</strong>'
    $html = $html -replace '\*([^\n\r*]+)\*', '<em>$1</em>'
    $html = $html -replace '~~([^\n\r~]+)~~', '<del>$1</del>'

    # Links and Images
    $html = $html -replace '!\[([^\]]*)\]\(([^)]+)\)', '<img src="$2" alt="$1"/>'
    $html = $html -replace '\[([^\]]+)\]\(([^)]+)\)', '<a href="$2" target="_blank" rel="noopener noreferrer">$1</a>'

    # Helper functions
    function Pop-Stack {
        param([ref]$Stack)
        if ($Stack.Value.Count -gt 0) {
            if ($Stack.Value.Count -eq 1) {
                $Stack.Value = @()  # Ensure it's an array
            }
            else {
                $Stack.Value = @($Stack.Value[0..($Stack.Value.Count - 2)])  # Ensure it's an array
            }
        }
    }

    function Update-ListNesting {
        param(
            [int]$TargetLevel,
            [ref]$ListStack,
            [ref]$ProcessedLines,
            [string]$ListType
        )

        $currentLevel = $ListStack.Value.Count

        if ($TargetLevel -gt $currentLevel) {
            for ($n = $currentLevel; $n -lt $TargetLevel; $n++) {
                $ProcessedLines.Value += "<$ListType>"
                $ListStack.Value += $ListType
            }
        }
        elseif ($TargetLevel -lt $currentLevel) {
            for ($n = $currentLevel; $n -gt $TargetLevel; $n--) {
                $closeType = $ListStack.Value[-1]
                $ProcessedLines.Value += "</$closeType>"
                Pop-Stack -Stack $ListStack
            }
        }
    }

    function Close-AllList {
        param(
            [ref]$ListStack,
            [ref]$ProcessedLines,
            [ref]$InUnorderedList,
            [ref]$InOrderedList
        )

        while ($ListStack.Value.Count -gt 0) {
            $listType = $ListStack.Value[-1]
            $closeTag = "</$listType>"
            $ProcessedLines.Value += $closeTag
            Pop-Stack -Stack $ListStack
        }
        $InUnorderedList.Value = $false
        $InOrderedList.Value = $false
    }

    # Single-pass line processing
    $lines = $html -split "`n"
    $processedLines = @()
    $lineCount = $lines.Count

    $inTable = $false
    $inUnorderedList = $false
    $inOrderedList = $false
    $inBlockquote = $false
    $tableAlignments = @()
    $listStack = @()

    for ($i = 0; $i -lt $lineCount; $i++) {
        $line = $lines[$i]

        # Blockquote processing
        if ($line -match '^>\s*(.*)$') {
            if ($inTable) { $processedLines += '</tbody></table></div>'; $inTable = $false; $tableAlignments = @() }
            Close-AllList -ListStack ([ref]$listStack) -ProcessedLines ([ref]$processedLines) -InUnorderedList ([ref]$inUnorderedList) -InOrderedList ([ref]$inOrderedList)

            $content = $Matches[1]
            if (-not $inBlockquote) {
                $processedLines += '<blockquote>'
                $inBlockquote = $true
            }
            if ($content.Trim() -ne '') {
                $processedLines += $content
            }
        }
        # Table processing
        elseif ($line -match '^\|.*\|$') {
            if ($inBlockquote) { $processedLines += '</blockquote>'; $inBlockquote = $false }
            Close-AllList -ListStack ([ref]$listStack) -ProcessedLines ([ref]$processedLines) -InUnorderedList ([ref]$inUnorderedList) -InOrderedList ([ref]$inOrderedList)

            if (-not $inTable) {
                $processedLines += '<div class="table-wrapper">'
                $processedLines += '<table class="table table-striped">'
                $inTable = $true

                # Check for separator line with alignment
                if (($i + 1) -lt $lineCount -and $lines[$i + 1] -match '^\|[-:\s\|]+\|$') {
                    $separatorLine = $lines[$i + 1]
                    $alignmentCells = ($separatorLine -replace '^\|', '' -replace '\|$', '').Split('|')
                    $tableAlignments = @()
                    foreach ($alignCell in $alignmentCells) {
                        $alignCell = $alignCell.Trim()
                        if ($alignCell -match '^:.*:$') { $tableAlignments += 'center' }
                        elseif ($alignCell -match ':$') { $tableAlignments += 'right' }
                        elseif ($alignCell -match '^:') { $tableAlignments += 'left' }
                        else { $tableAlignments += '' }
                    }

                    # Process header row
                    $tempLine = $line -replace '\\\|', '§PIPE§'
                    $cells = ($tempLine -replace '^\|', '' -replace '\|$', '').Split('|')
                    if ($cells.Count -gt 0) {
                        $processedLines += '<thead><tr>'
                        for ($j = 0; $j -lt $cells.Count; $j++) {
                            $cleanCell = $cells[$j].Trim() -replace '§PIPE§', '|'
                            if ([string]::IsNullOrWhiteSpace($cleanCell)) { $cleanCell = '&nbsp;' }
                            $alignClass = if ($j -lt $tableAlignments.Count -and $tableAlignments[$j]) { " class=`"text-$($tableAlignments[$j])`"" } else { "" }
                            $processedLines += "<th$alignClass>$cleanCell</th>"
                        }
                        $processedLines += '</tr></thead><tbody>'
                        $i++
                        continue
                    }
                }
            }

            # Regular table row
            $tempLine = $line -replace '\\\|', '§PIPE§'
            $cells = ($tempLine -replace '^\|', '' -replace '\|$', '').Split('|')
            if ($cells.Count -gt 0) {
                $processedLines += '<tr>'
                for ($j = 0; $j -lt $cells.Count; $j++) {
                    $cleanCell = $cells[$j].Trim() -replace '§PIPE§', '|'
                    if ([string]::IsNullOrWhiteSpace($cleanCell)) { $cleanCell = '&nbsp;' }
                    $alignClass = if ($j -lt $tableAlignments.Count -and $tableAlignments[$j]) { " class=`"text-$($tableAlignments[$j])`"" } else { "" }
                    $processedLines += "<td$alignClass>$cleanCell</td>"
                }
                $processedLines += '</tr>'
            }
        }
        # Unordered List processing
        elseif ($line -match '^(\s*)- (.+)$') {
            if ($inBlockquote) { $processedLines += '</blockquote>'; $inBlockquote = $false }
            if ($inTable) { $processedLines += '</tbody></table></div>'; $inTable = $false; $tableAlignments = @() }
            if ($inOrderedList) { $processedLines += '</ol>'; $inOrderedList = $false }

            $indentation = $Matches[1].Length
            $content = $Matches[2]
            $nestLevel = [Math]::Floor($indentation / 2)

            # Open first list if needed
            if (-not $inUnorderedList) {
                $processedLines += '<ul>'
                $inUnorderedList = $true
                $listStack += 'ul'
            }

            # Handle nesting (nestLevel+1 because nestLevel is 0-based, only update if different)
            $targetLevel = $nestLevel + 1
            if ($targetLevel -ne $listStack.Count) {
                Update-ListNesting -TargetLevel $targetLevel -ListStack ([ref]$listStack) -ProcessedLines ([ref]$processedLines) -ListType 'ul'
            }

            $processedLines += "<li>$content</li>"
        }
        # Ordered List processing
        elseif ($line -match '^(\s*)(\d+)\. (.+)$') {
            if ($inBlockquote) { $processedLines += '</blockquote>'; $inBlockquote = $false }
            if ($inTable) { $processedLines += '</tbody></table></div>'; $inTable = $false; $tableAlignments = @() }
            if ($inUnorderedList) { $processedLines += '</ul>'; $inUnorderedList = $false }

            $indentation = $Matches[1].Length
            $content = $Matches[3]
            $nestLevel = [Math]::Floor($indentation / 2)

            # Open first list if needed
            if (-not $inOrderedList) {
                $processedLines += '<ol>'
                $inOrderedList = $true
                $listStack += 'ol'
            }

            # Handle nesting (nestLevel+1 because nestLevel is 0-based, only update if different)
            $targetLevel = $nestLevel + 1
            if ($targetLevel -ne $listStack.Count) {
                Update-ListNesting -TargetLevel $targetLevel -ListStack ([ref]$listStack) -ProcessedLines ([ref]$processedLines) -ListType 'ol'
            }

            $processedLines += "<li>$content</li>"
        }
        # Other lines
        else {
            if ($inBlockquote) { $processedLines += '</blockquote>'; $inBlockquote = $false }
            if ($inTable) { $processedLines += '</tbody></table></div>'; $inTable = $false; $tableAlignments = @() }

            $isHeader = $line -match '^<h[1-6]>'
            $isEmptyLine = [string]::IsNullOrWhiteSpace($line)
            $nextLineIsList = $false
            $nextLineIsHeader = $false

            if ($isEmptyLine -and ($i + 1) -lt $lineCount) {
                for ($j = $i + 1; $j -lt $lineCount; $j++) {
                    $nextLine = $lines[$j]
                    if (-not [string]::IsNullOrWhiteSpace($nextLine)) {
                        $nextLineIsList = ($nextLine -match '^(\s*)- (.+)$') -or ($nextLine -match '^(\s*)(\d+)\. (.+)$')
                        $nextLineIsHeader = ($nextLine -match '^<h[1-6]>')
                        break
                    }
                }
            }

            if ($listStack.Count -gt 0 -and ($isHeader -or ($isEmptyLine -and -not $nextLineIsList -and -not $nextLineIsHeader))) {
                Close-AllList -ListStack ([ref]$listStack) -ProcessedLines ([ref]$processedLines) -InUnorderedList ([ref]$inUnorderedList) -InOrderedList ([ref]$inOrderedList)
            }

            if (-not $isEmptyLine -or $listStack.Count -eq 0) {
                $processedLines += $line
            }
        }
    }

    # Close remaining open structures
    if ($inBlockquote) { $processedLines += '</blockquote>' }
    if ($inTable) { $processedLines += '</tbody></table></div>' }
    Close-AllList -ListStack ([ref]$listStack) -ProcessedLines ([ref]$processedLines) -InUnorderedList ([ref]$inUnorderedList) -InOrderedList ([ref]$inOrderedList)

    $html = $processedLines -join "`n"

    # Paragraph processing
    $blocks = $html -split "`n`n+"

    $result = @()
    foreach ($block in $blocks) {
        $block = $block.Trim()
        if ($block -eq "") { continue }

        # Check if block starts with an HTML element tag (opening or closing)
        if ($block -match "^<(h[1-6]|ul|ol|table|pre|blockquote|hr)[\s>]" -or
            $block -match "^</(h[1-6]|ul|ol|table|pre|blockquote)>") {
            $result += $block
        }
        # Check if it contains HTML list elements - if so, don't wrap
        elseif ($block -match "<(h[1-6]|ul|ol|li|table|thead|tbody|tr|td|th|pre|code|blockquote|hr|/ul|/ol)[\s>]") {
            $result += $block
        }
        else {
            $lines = $block -split "`n"
            $nonEmptyLines = @($lines | Where-Object { $_.Trim() -ne "" })
            if ($nonEmptyLines.Count -gt 0) {
                $paragraphContent = $nonEmptyLines -join '<br>'
                $result += "<p>$paragraphContent</p>"
            }
        }
    }

    $html = $result -join "`n`n"

    # Final safety escaping
    $html = $html -replace '&(?![a-zA-Z]{2,8};)(?!#[0-9]{1,7};)(?!#x[0-9a-fA-F]{1,6};)', '&amp;'

    # Restore escaped Markdown characters
    $html = $html -replace '§ESCAPED§(.{1})§ESCAPED§', '$1'

    # Restore inline code blocks from placeholders
    for ($i = 0; $i -lt $inlineCodeBlocks.Count; $i++) {
        $html = $html -replace "§INLINECODE§$i§", $inlineCodeBlocks[$i]
    }

    # Restore code blocks from placeholders
    for ($i = 0; $i -lt $codeBlocks.Count; $i++) {
        $html = $html -replace "§CODEBLOCK§$i§", $codeBlocks[$i]
    }

    return $html
}

function Get-RjReportEmailBody {
    <#
        .SYNOPSIS
        Builds the RealmJoin-branded HTML email body used for report delivery.

        .DESCRIPTION
        Assembles the static HTML template, injects the converted Markdown content, and renders
        optional attachment metadata as well as tenant information into the footer section.

        .PARAMETER Subject
        Subject of the email, used for the HTML <title> element.

        .PARAMETER HtmlContent
        HTML fragment generated from Markdown that will be embedded in the email body.

        .PARAMETER Attachments
        Optional list of attachment file paths to surface in the "Attached Files" section.

        .PARAMETER TenantDisplayName
        Optional tenant display name shown in the tenant information box.

        .PARAMETER ReportVersion
        Optional report version string rendered in the tenant information box.

        .OUTPUTS
        System.String. Returns the composed HTML email body.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$HtmlContent,

        [string[]]$Attachments = @(),

        [string]$TenantDisplayName,

        [string]$ReportVersion
    )

    if (-not $Attachments) {
        $Attachments = @()
    }

    $svgRJHeader = @'
<svg width="750" height="200" viewBox="0 0 750 200" fill="none" xmlns="http://www.w3.org/2000/svg">
<mask id="mask0_2027_657" style="mask-type:luminance" maskUnits="userSpaceOnUse" x="0" y="0" width="750" height="200">
<path d="M750 0H0V199.273H750V0Z" fill="white"/>
</mask>
<g mask="url(#mask0_2027_657)">
<path d="M750 0H0V223.721H750V0Z" fill="#011E33"/>
<path d="M580.133 -26.718L553.021 385.171L870.344 449.015V131.863L580.133 -26.718Z" fill="white" fill-opacity="0.03"/>
<path d="M0 152.6L64.5848 419.07L340.064 474.496L556.104 17.7625L0 152.6Z" fill="url(#paint0_linear_2027_657)" fill-opacity="0.1"/>
<g opacity="0.18">
<path d="M33.249 -2.16269L155.367 -29.3821L791.005 -20.5171L757.762 223.124L33.249 -2.16269Z" fill="#103753"/>
</g>
</g>
<g clip-path="url(#clip0_2027_657)">
<path d="M78.062 68.3951C77.9814 68.3129 77.9814 68.2307 77.9814 68.1485V52.7729C77.9814 52.5262 78.062 52.3618 78.223 52.1973C78.3841 52.0329 78.6257 51.9507 78.7867 51.9507H82.8131C84.5042 51.9507 86.1148 52.8551 86.9201 54.3351C87.3227 55.0751 87.5643 55.8974 87.5643 56.7196C87.5643 57.8707 87.2422 58.9396 86.5174 59.7618C85.7927 60.584 84.9069 61.1596 83.7795 61.4062L87.6448 67.8196C87.7254 67.9018 87.7254 67.984 87.7254 68.0662C87.7254 68.1485 87.6448 68.3129 87.5643 68.3951C87.4838 68.4774 87.4033 68.4774 87.2422 68.4774H86.0343C85.8732 68.4774 85.7122 68.3951 85.7122 68.2307L81.6052 61.4885H79.8336V68.1485C79.8336 68.2307 79.8336 68.3129 79.7531 68.3951C79.6725 68.4774 79.592 68.4774 79.5115 68.4774H78.3036C78.223 68.4774 78.062 68.3951 78.062 68.3951ZM82.7326 59.6796C83.5379 59.6796 84.2627 59.4329 84.8264 58.8574C85.3901 58.2818 85.6316 57.5418 85.6316 56.8018C85.6316 55.9796 85.3901 55.2396 84.8264 54.664C84.2627 54.0885 83.5379 53.8418 82.7326 53.8418H79.8336V59.7618L82.7326 59.6796Z" fill="white"/>
<path d="M91.8323 67.7372C91.027 67.3261 90.4633 66.6683 89.9802 65.8461C89.497 65.0239 89.3359 64.1194 89.3359 63.215V59.7617C89.3359 58.8572 89.5775 57.9528 89.9802 57.1306C90.866 55.4861 92.5571 54.4994 94.3287 54.5817C95.2145 54.5817 96.1003 54.8283 96.9056 55.2394C97.7109 55.6506 98.2746 56.3083 98.7578 57.1306C99.2409 57.9528 99.402 58.8572 99.402 59.7617V61.4061C99.402 61.6528 99.3215 61.8172 99.1604 61.9817C98.9993 62.1461 98.7578 62.2283 98.5967 62.2283H91.1076V63.215C91.1076 63.7906 91.2686 64.3661 91.5102 64.8594C91.7518 65.3528 92.235 65.7639 92.7181 66.0928C93.2013 66.3394 93.765 66.5039 94.3287 66.5039H97.7109C97.7914 66.5039 97.8719 66.5039 97.9525 66.5861C98.033 66.6683 98.033 66.7506 98.033 66.8328V67.9017C98.033 68.0661 97.8719 68.2306 97.7109 68.2306H94.3287C93.4429 68.4772 92.5571 68.2306 91.8323 67.7372ZM97.6304 60.5839V59.7617C97.6304 59.1861 97.4693 58.6106 97.2277 58.1172C96.9861 57.6239 96.5835 57.2128 96.0198 56.8839C94.4898 55.9794 92.4765 56.4728 91.5907 58.035V58.1172C91.2686 58.6106 91.1076 59.1861 91.1076 59.7617V60.5839H97.6304Z" fill="white"/>
<path d="M101.978 67.7374C100.529 66.3396 100.448 63.9552 101.817 62.393C102.622 61.653 103.669 61.2418 104.716 61.0774L107.776 60.6663C108.179 60.6663 108.501 60.5018 108.743 60.173C108.984 59.8441 109.065 59.5152 109.065 59.1041V58.9396C109.065 58.1996 108.743 57.5418 108.179 57.1307C107.535 56.6374 106.81 56.3907 106.005 56.3907C105.36 56.3907 104.797 56.5552 104.233 56.8018C103.75 57.0485 103.347 57.4596 103.025 57.953C102.945 58.1174 102.784 58.2818 102.542 58.2818C102.461 58.2818 102.3 58.1996 102.22 58.1996L101.576 57.7885C101.415 57.7063 101.334 57.5418 101.334 57.3774C101.334 57.2952 101.334 57.213 101.415 57.1307C101.898 56.3907 102.542 55.733 103.267 55.3218C104.072 54.8285 105.038 54.6641 106.005 54.6641C106.89 54.6641 107.696 54.8285 108.501 55.1574C109.226 55.4863 109.79 55.9796 110.192 56.6374C110.595 57.2952 110.836 58.1174 110.756 58.9396V68.1485C110.756 68.313 110.595 68.4774 110.434 68.4774H109.467C109.387 68.4774 109.306 68.4774 109.226 68.3952C109.145 68.313 109.145 68.2307 109.145 68.1485V66.5041C108.662 67.1618 108.018 67.7374 107.293 68.1485C106.568 68.5596 105.683 68.7241 104.797 68.7241C103.83 68.8063 102.784 68.3952 101.978 67.7374ZM107.132 66.5041C107.776 66.093 108.259 65.5996 108.582 64.9418C108.904 64.3663 109.065 63.7085 109.145 63.0507V62.0641L105.522 62.5574L104.877 62.6396C103.347 62.8863 102.542 63.6263 102.542 64.8596C102.542 65.4352 102.784 66.0107 103.267 66.4218C103.75 66.833 104.394 67.0796 105.119 66.9974C105.844 67.0796 106.568 66.9152 107.132 66.5041Z" fill="white"/>
<path d="M114.863 68.2306C114.702 68.0662 114.621 67.8195 114.621 67.6551V52.1151H112.447C112.367 52.1151 112.286 52.1151 112.206 52.0329C112.125 51.9506 112.125 51.8684 112.125 51.7862V50.7173C112.125 50.6351 112.125 50.5529 112.206 50.4707C112.286 50.3884 112.367 50.3884 112.447 50.3884H115.588C115.829 50.3884 115.99 50.4706 116.151 50.6351C116.312 50.7995 116.393 51.0462 116.393 51.2106V66.6684H118.889C118.97 66.6684 119.05 66.6684 119.131 66.7507C119.211 66.8329 119.211 66.9151 119.211 66.9973V68.0662C119.211 68.2306 119.05 68.3951 118.889 68.3951H115.427C115.266 68.4773 115.024 68.3951 114.863 68.2306Z" fill="white"/>
<path d="M121.225 68.3952C121.145 68.313 121.145 68.2307 121.145 68.1485V55.3218C121.145 55.2396 121.145 55.1574 121.225 55.0752C121.306 54.993 121.386 54.993 121.467 54.993H122.594C122.675 54.993 122.755 54.993 122.836 55.0752C122.916 55.1574 122.916 55.2396 122.916 55.3218V56.6374C123.238 56.0618 123.802 55.4863 124.366 55.1574C125.01 54.8285 125.735 54.6641 126.379 54.6641C127.265 54.6641 128.15 54.9107 128.956 55.4041C129.681 55.8974 130.164 56.5552 130.486 57.3774C130.808 56.5552 131.291 55.8974 132.016 55.4041C132.741 54.993 133.546 54.7463 134.351 54.7463C135.559 54.6641 136.686 55.1574 137.572 55.9796C138.378 56.8841 138.861 58.0352 138.78 59.2685V68.1485C138.78 68.313 138.619 68.4774 138.458 68.4774H137.331C137.17 68.4774 137.009 68.313 137.009 68.1485V59.5152C137.089 58.693 136.767 57.7885 136.203 57.213C135.64 56.6374 134.834 56.3907 134.11 56.3907C133.224 56.3907 132.338 56.7196 131.774 57.4596C131.13 58.2818 130.808 59.2685 130.888 60.3374V68.1485C130.888 68.313 130.727 68.4774 130.566 68.4774H129.439C129.278 68.4774 129.117 68.313 129.117 68.1485V59.5152C129.197 58.693 128.875 57.7885 128.312 57.213C127.748 56.6374 127.023 56.3907 126.218 56.3907C125.332 56.3907 124.446 56.7196 123.882 57.4596C123.238 58.2818 122.916 59.2685 122.997 60.3374V68.1485C122.997 68.2307 122.997 68.313 122.916 68.3952C122.836 68.4774 122.755 68.4774 122.675 68.4774H121.547C121.386 68.4226 121.279 68.3952 121.225 68.3952Z" fill="white"/>
<path d="M140.471 68.3951C140.391 68.3129 140.391 68.2306 140.391 68.1484V67.0795C140.391 66.9151 140.471 66.7506 140.632 66.7506H142.082C142.565 66.7506 143.048 66.5862 143.37 66.2573C143.692 65.8462 143.934 65.3529 143.853 64.8595V52.1973C143.853 52.1151 143.853 52.0329 143.934 51.9506C144.014 51.8684 144.095 51.8684 144.175 51.8684H145.464C145.544 51.8684 145.625 51.8684 145.706 51.9506C145.786 52.0329 145.786 52.1151 145.786 52.1973V64.6951C145.786 65.3529 145.625 66.0106 145.383 66.5862C145.142 67.1617 144.739 67.6551 144.175 67.984C143.612 68.3129 142.968 68.4773 142.323 68.4773H140.713C140.632 68.4773 140.552 68.4773 140.471 68.3951Z" fill="white"/>
<path d="M150.699 68.0662C149.893 67.5729 149.33 66.9973 148.847 66.1751C148.363 65.3529 148.202 64.4484 148.202 63.544V59.844C148.122 56.9662 150.296 54.5817 153.034 54.4995C155.853 54.4173 158.188 56.6373 158.268 59.4329V63.544C158.268 64.4484 158.027 65.3529 157.624 66.1751C156.255 68.6417 153.115 69.5462 150.699 68.0662ZM154.886 66.5862C155.369 66.3395 155.772 65.8462 156.094 65.3529C156.416 64.8595 156.497 64.284 156.497 63.7084V59.7617C156.497 59.1862 156.336 58.6106 156.094 58.1173C155.853 57.624 155.369 57.2129 154.886 56.884C154.403 56.6373 153.839 56.4729 153.276 56.4729C152.712 56.4729 152.148 56.6373 151.665 56.884C151.182 57.2129 150.779 57.624 150.457 58.1173C150.135 58.6106 150.055 59.1862 150.055 59.7617V63.7084C150.055 64.284 150.216 64.8595 150.457 65.3529C150.699 65.8462 151.182 66.2573 151.665 66.5862C152.148 66.9151 152.712 66.9973 153.276 66.9973C153.839 66.9973 154.403 66.8329 154.886 66.5862Z" fill="white"/>
<path d="M162.294 68.2306C162.133 68.0662 162.052 67.8195 162.052 67.6551V56.7195H159.637C159.476 56.7195 159.314 56.6373 159.314 56.3906V55.3218C159.314 55.1573 159.395 54.9929 159.637 54.9929H163.019C163.26 54.9929 163.421 55.0751 163.582 55.2395C163.744 55.404 163.824 55.6507 163.824 55.8151V66.6684H166.24C166.401 66.6684 166.562 66.8329 166.562 66.9973V68.0662C166.562 68.1484 166.562 68.2307 166.481 68.3129C166.401 68.3951 166.32 68.3951 166.24 68.3951H162.858C162.616 68.4773 162.455 68.3951 162.294 68.2306ZM162.133 53.0195C162.052 52.9373 162.052 52.8551 162.052 52.7729V50.7173C162.052 50.6351 162.052 50.5529 162.133 50.4707C162.213 50.3884 162.294 50.3884 162.375 50.3884H163.502C163.582 50.3884 163.663 50.3884 163.744 50.4707C163.824 50.5529 163.824 50.6351 163.824 50.7173V52.7729C163.824 52.8551 163.824 52.9373 163.744 53.0195C163.663 53.1018 163.582 53.1018 163.502 53.1018H162.375C162.294 53.1018 162.213 53.1018 162.133 53.0195Z" fill="white"/>
<path d="M168.657 68.3128C168.576 68.2306 168.576 68.1484 168.576 68.0662V55.3217C168.576 55.1573 168.737 54.9928 168.898 54.9928H170.026C170.187 54.9928 170.348 55.075 170.348 55.3217V56.6373C170.75 55.9795 171.314 55.4862 171.958 55.1573C172.603 54.8284 173.408 54.6639 174.133 54.6639C175.341 54.5817 176.468 55.075 177.354 55.8973C178.159 56.8017 178.642 57.9528 178.562 59.1862V68.1484C178.562 68.3128 178.481 68.4773 178.24 68.4773H177.112C176.951 68.4773 176.79 68.3128 176.79 68.1484V59.515C176.871 58.6928 176.548 57.7884 175.985 57.2128C175.421 56.6373 174.616 56.3906 173.891 56.3906C172.925 56.3906 172.039 56.7195 171.395 57.4595C170.75 58.2817 170.348 59.2684 170.428 60.3373V68.1484C170.428 68.3128 170.348 68.4773 170.106 68.4773H168.979C168.818 68.395 168.737 68.395 168.657 68.3128Z" fill="white"/>
<path d="M56.642 65.9285C56.642 65.5996 56.8835 65.2707 57.2057 65.2707H59.541C60.0242 65.2707 60.4268 65.1063 60.7489 64.7774C61.071 64.4485 61.2321 63.9552 61.1516 63.4618V47.1818C61.1516 46.8529 61.3931 46.5241 61.7958 46.5241H64.2116C64.5337 46.5241 64.8559 46.7707 64.8559 47.0996V63.4618C64.8559 64.4485 64.6143 65.353 64.2116 66.2574C63.809 67.0796 63.1648 67.7374 62.3595 68.1485C61.5542 68.6418 60.6684 68.8885 59.702 68.8885H57.9304L70.976 71.2729L67.1107 45.7841L43.4353 34.2729L41.9053 46.7707C42.1469 46.6063 42.3884 46.5241 42.63 46.5241H48.3475C49.5555 46.5241 50.7634 46.7707 51.7297 47.4285C52.7766 48.0041 53.5819 48.8263 54.2261 49.8952C55.5951 52.1974 55.434 55.1574 53.8235 57.2952C52.9377 58.4463 51.7297 59.3507 50.3608 59.7618L55.434 67.8196C55.5146 67.9841 55.5146 68.0663 55.5951 68.2307C55.5951 68.313 55.5951 68.3952 55.5146 68.4774L56.7225 68.7241C56.642 68.6418 56.5614 68.4774 56.5614 68.313V65.9285H56.642Z" fill="#F8842C"/>
<path d="M59.782 68.8885C60.6678 68.8885 61.6342 68.6418 62.4395 68.1485C63.2447 67.6551 63.8084 66.9973 64.2916 66.2573C64.6943 65.4351 64.9358 64.4485 64.9358 63.4618V47.0996C64.9358 46.7707 64.6137 46.4418 64.2916 46.524H61.8758C61.5537 46.524 61.2315 46.8529 61.2315 47.1818V63.4618C61.2315 63.9551 61.0705 64.4485 60.8289 64.7773C60.5068 65.1062 60.1041 65.2707 59.621 65.2707H57.2051C56.883 65.2707 56.5609 65.5996 56.6414 65.9285V68.2307C56.6414 68.3951 56.7219 68.5596 56.8025 68.6418L58.0104 68.8885H59.782Z" fill="white"/>
<path d="M45.0449 60.0906V66.5039L51.1651 67.655L46.575 60.0906H45.0449Z" fill="#F8842C"/>
<path d="M39.5693 65.5174L41.341 65.8462V51.2107L39.5693 65.5174Z" fill="#F8842C"/>
<path d="M51.4067 53.3485C51.4067 52.4441 51.0845 51.6218 50.5209 51.0463C49.8766 50.4707 49.0713 50.1418 48.1855 50.1418H45.0449V56.5552H48.1855C49.0713 56.5552 49.8766 56.3085 50.5209 55.6507C51.0845 55.0752 51.4067 54.253 51.4067 53.3485Z" fill="#F8842C"/>
<path d="M55.6749 68.2308C55.6749 68.0664 55.5943 67.9019 55.5138 67.8197L50.4405 59.7619C51.8095 59.3508 53.0174 58.4464 53.9032 57.2953C55.5138 55.1575 55.6749 52.1975 54.3059 49.8953C53.7422 48.8264 52.8564 48.0042 51.8095 47.4286C50.7626 46.8531 49.5547 46.5242 48.4273 46.5242H42.6293C42.3877 46.5242 42.0656 46.6064 41.9045 46.7708L41.3408 51.2108V65.8464L45.0451 66.5042V60.0908H46.5752L51.1653 67.6553L55.5943 68.4775C55.5943 68.3953 55.6749 68.3131 55.6749 68.2308ZM45.0451 56.5553V50.1419H48.1857C49.0715 50.1419 49.8768 50.4708 50.5211 51.0464C51.0848 51.6219 51.4874 52.5264 51.4069 53.3486C51.4069 54.1708 51.0848 54.9931 50.5211 55.6508C49.8768 56.2264 49.0715 56.5553 48.1857 56.5553H45.0451Z" fill="white"/>
</g>
<g filter="url(#filter0_d_2027_657)">
<rect x="201.548" y="117.351" width="346.905" height="11.2236" fill="#F8842C"/>
</g>
<path d="M213.773 122.561C213.636 122.561 213.521 122.515 213.43 122.424C213.338 122.332 213.292 122.218 213.292 122.08V120.466C213.292 120.329 213.338 120.214 213.43 120.123C213.521 120.031 213.636 119.985 213.773 119.985H217.242V101.268H213.773C213.636 101.268 213.521 101.222 213.43 101.13C213.338 101.039 213.292 100.924 213.292 100.787V99.1727C213.292 99.0353 213.338 98.9208 213.43 98.8292C213.521 98.7377 213.636 98.6919 213.773 98.6919H223.527C223.664 98.6919 223.779 98.7377 223.87 98.8292C223.962 98.9208 224.008 99.0353 224.008 99.1727V100.787C224.008 100.924 223.962 101.039 223.87 101.13C223.779 101.222 223.664 101.268 223.527 101.268H220.058V119.985H223.527C223.664 119.985 223.779 120.031 223.87 120.123C223.962 120.214 224.008 120.329 224.008 120.466V122.08C224.008 122.218 223.962 122.332 223.87 122.424C223.779 122.515 223.664 122.561 223.527 122.561H213.773ZM229.825 122.527C229.688 122.527 229.573 122.481 229.482 122.389C229.39 122.298 229.344 122.183 229.344 122.046V103.603C229.344 103.466 229.39 103.351 229.482 103.26C229.573 103.168 229.688 103.122 229.825 103.122H231.542C231.68 103.122 231.794 103.168 231.886 103.26C231.977 103.351 232.023 103.466 232.023 103.603V105.492C232.618 104.553 233.385 103.843 234.324 103.363C235.286 102.882 236.362 102.641 237.553 102.641C239.499 102.641 241.067 103.237 242.258 104.427C243.471 105.618 244.078 107.186 244.078 109.132V122.046C244.078 122.183 244.032 122.298 243.941 122.389C243.849 122.481 243.735 122.527 243.597 122.527H241.88C241.743 122.527 241.628 122.481 241.536 122.389C241.445 122.298 241.399 122.183 241.399 122.046V109.648C241.399 108.182 241.01 107.072 240.231 106.316C239.476 105.538 238.446 105.149 237.14 105.149C235.606 105.149 234.37 105.652 233.431 106.66C232.493 107.667 232.023 109.052 232.023 110.815V122.046C232.023 122.183 231.977 122.298 231.886 122.389C231.794 122.481 231.68 122.527 231.542 122.527H229.825ZM254.575 122.904C253.521 122.904 252.491 122.767 251.484 122.492C250.476 122.218 249.606 121.817 248.874 121.29C248.782 121.245 248.713 121.176 248.667 121.084C248.645 120.97 248.656 120.855 248.702 120.741L249.423 119.298C249.537 119.115 249.686 119.024 249.869 119.024C249.915 119.024 250.007 119.046 250.144 119.092C251.655 120.008 253.178 120.466 254.712 120.466C255.926 120.466 256.876 120.203 257.563 119.676C258.272 119.127 258.627 118.36 258.627 117.375C258.627 116.78 258.467 116.299 258.146 115.933C257.849 115.543 257.448 115.234 256.944 115.005C256.441 114.753 255.651 114.421 254.575 114.009C253.178 113.483 252.079 112.99 251.278 112.533C250.476 112.075 249.847 111.491 249.389 110.781C248.931 110.071 248.702 109.167 248.702 108.068C248.702 106.305 249.286 104.965 250.453 104.05C251.621 103.134 253.121 102.676 254.952 102.676C255.914 102.676 256.864 102.813 257.803 103.088C258.765 103.363 259.623 103.752 260.379 104.256C260.47 104.301 260.528 104.382 260.551 104.496C260.573 104.611 260.562 104.714 260.516 104.805L259.829 106.213C259.783 106.351 259.692 106.442 259.555 106.488C259.44 106.534 259.314 106.511 259.177 106.419C257.711 105.549 256.303 105.114 254.952 105.114C253.831 105.114 252.949 105.355 252.308 105.835C251.69 106.316 251.381 107.015 251.381 107.93C251.381 108.64 251.552 109.213 251.896 109.648C252.239 110.083 252.651 110.415 253.132 110.644C253.613 110.85 254.38 111.147 255.433 111.537C256.784 112.017 257.849 112.464 258.627 112.876C259.406 113.288 260.047 113.838 260.551 114.525C261.054 115.211 261.306 116.104 261.306 117.203C261.306 118.966 260.688 120.363 259.452 121.393C258.238 122.401 256.612 122.904 254.575 122.904ZM269.058 100.443C268.921 100.443 268.807 100.398 268.715 100.306C268.623 100.214 268.578 100.1 268.578 99.9626V96.9747C268.578 96.8373 268.623 96.7228 268.715 96.6312C268.807 96.5396 268.921 96.4938 269.058 96.4938H270.707C270.844 96.4938 270.959 96.5396 271.05 96.6312C271.142 96.7228 271.188 96.8373 271.188 96.9747V99.9626C271.188 100.1 271.142 100.214 271.05 100.306C270.959 100.398 270.844 100.443 270.707 100.443H269.058ZM269.78 122.561C269.436 122.561 269.139 122.447 268.887 122.218C268.658 121.966 268.543 121.668 268.543 121.325V105.664H264.972C264.834 105.664 264.72 105.618 264.628 105.526C264.536 105.435 264.491 105.32 264.491 105.183V103.603C264.491 103.466 264.536 103.351 264.628 103.26C264.72 103.168 264.834 103.122 264.972 103.122H269.986C270.329 103.122 270.615 103.248 270.844 103.5C271.096 103.729 271.222 104.015 271.222 104.359V120.02H274.76C274.897 120.02 275.011 120.065 275.103 120.157C275.195 120.249 275.24 120.363 275.24 120.5V122.08C275.24 122.218 275.195 122.332 275.103 122.424C275.011 122.515 274.897 122.561 274.76 122.561H269.78ZM285.842 129.636C283.667 129.636 281.893 129.235 280.519 128.434C279.168 127.633 278.493 126.442 278.493 124.862C278.493 124.061 278.676 123.294 279.042 122.561C279.408 121.828 279.901 121.222 280.519 120.741C279.947 120.352 279.523 119.962 279.248 119.573C278.996 119.161 278.87 118.669 278.87 118.096C278.87 117.387 279.065 116.711 279.454 116.07C279.866 115.429 280.416 114.891 281.103 114.456C280.324 113.861 279.718 113.128 279.283 112.258C278.848 111.365 278.63 110.38 278.63 109.304C278.63 107.999 278.939 106.843 279.557 105.835C280.198 104.828 281.057 104.05 282.133 103.5C283.209 102.951 284.4 102.676 285.705 102.676C286.598 102.676 287.491 102.825 288.384 103.122H295.184C295.321 103.122 295.436 103.168 295.527 103.26C295.619 103.351 295.665 103.466 295.665 103.603V105.183C295.665 105.32 295.619 105.435 295.527 105.526C295.436 105.618 295.321 105.664 295.184 105.664H291.715C292.425 106.694 292.78 107.908 292.78 109.304C292.78 110.609 292.459 111.766 291.818 112.773C291.2 113.78 290.353 114.559 289.277 115.108C288.201 115.658 287.01 115.933 285.705 115.933C284.835 115.933 283.93 115.784 282.992 115.486C282.488 115.715 282.099 116.013 281.824 116.379C281.549 116.745 281.412 117.135 281.412 117.547C281.412 118.073 281.641 118.463 282.099 118.714C282.465 118.921 283.003 119.069 283.713 119.161C284.423 119.253 285.533 119.367 287.044 119.504L288.384 119.607C289.483 119.699 290.421 119.848 291.2 120.054C292.001 120.26 292.677 120.661 293.226 121.256C293.799 121.851 294.085 122.721 294.085 123.866C294.085 125.744 293.318 127.175 291.784 128.159C290.25 129.144 288.269 129.636 285.842 129.636ZM285.705 113.563C287.056 113.563 288.132 113.174 288.933 112.395C289.758 111.594 290.17 110.563 290.17 109.304C290.17 108.045 289.758 107.026 288.933 106.248C288.132 105.446 287.056 105.046 285.705 105.046C284.354 105.046 283.266 105.446 282.442 106.248C281.641 107.026 281.24 108.045 281.24 109.304C281.24 110.563 281.641 111.594 282.442 112.395C283.266 113.174 284.354 113.563 285.705 113.563ZM286.151 127.266C287.754 127.266 289.059 127.014 290.067 126.511C291.074 126.007 291.578 125.217 291.578 124.141C291.578 123.591 291.429 123.179 291.131 122.904C290.857 122.63 290.467 122.447 289.964 122.355C289.483 122.24 288.762 122.137 287.8 122.046C287.594 122.023 287.102 121.989 286.323 121.943C285.568 121.874 284.732 121.783 283.816 121.668C283.266 121.599 282.809 121.531 282.442 121.462C281.984 121.805 281.618 122.229 281.343 122.733C281.091 123.236 280.965 123.752 280.965 124.278C280.965 125.331 281.446 126.087 282.408 126.545C283.392 127.026 284.64 127.266 286.151 127.266ZM299.922 122.561C299.785 122.561 299.67 122.515 299.579 122.424C299.487 122.332 299.441 122.218 299.441 122.08V96.9747C299.441 96.8373 299.487 96.7228 299.579 96.6312C299.67 96.5396 299.785 96.4938 299.922 96.4938H301.639C301.777 96.4938 301.891 96.5396 301.983 96.6312C302.074 96.7228 302.12 96.8373 302.12 96.9747V105.526C302.715 104.588 303.482 103.878 304.421 103.397C305.383 102.916 306.459 102.676 307.649 102.676C309.596 102.676 311.164 103.271 312.355 104.462C313.568 105.652 314.175 107.221 314.175 109.167V122.08C314.175 122.218 314.129 122.332 314.037 122.424C313.946 122.515 313.831 122.561 313.694 122.561H311.977C311.839 122.561 311.725 122.515 311.633 122.424C311.542 122.332 311.496 122.218 311.496 122.08V109.682C311.496 108.217 311.107 107.106 310.328 106.351C309.573 105.572 308.542 105.183 307.237 105.183C305.703 105.183 304.467 105.687 303.528 106.694C302.589 107.701 302.12 109.087 302.12 110.85V122.08C302.12 122.218 302.074 122.332 301.983 122.424C301.891 122.515 301.777 122.561 301.639 122.561H299.922ZM326.152 122.561C324.504 122.561 323.21 122.069 322.271 121.084C321.333 120.077 320.863 118.818 320.863 117.306V105.629H318.597C318.459 105.629 318.345 105.584 318.253 105.492C318.162 105.4 318.116 105.286 318.116 105.149V103.603C318.116 103.466 318.162 103.351 318.253 103.26C318.345 103.168 318.459 103.122 318.597 103.122H320.863V99.5161C320.863 99.3787 320.909 99.2643 321.001 99.1727C321.092 99.0811 321.207 99.0353 321.344 99.0353H323.061C323.199 99.0353 323.313 99.0811 323.405 99.1727C323.496 99.2643 323.542 99.3787 323.542 99.5161V103.122H329.587C329.724 103.122 329.839 103.168 329.93 103.26C330.022 103.351 330.068 103.466 330.068 103.603V105.183C330.068 105.32 330.022 105.435 329.93 105.526C329.839 105.618 329.724 105.664 329.587 105.664H323.542V116.997C323.542 117.89 323.806 118.611 324.332 119.161C324.882 119.688 325.591 119.962 326.461 119.985L329.587 120.02C329.724 120.02 329.839 120.065 329.93 120.157C330.022 120.249 330.068 120.363 330.068 120.5V122.08C330.068 122.218 330.022 122.332 329.93 122.424C329.839 122.515 329.724 122.561 329.587 122.561H326.152ZM339.731 122.904C338.677 122.904 337.647 122.767 336.64 122.492C335.632 122.218 334.762 121.817 334.03 121.29C333.938 121.245 333.869 121.176 333.823 121.084C333.801 120.97 333.812 120.855 333.858 120.741L334.579 119.298C334.694 119.115 334.842 119.024 335.025 119.024C335.071 119.024 335.163 119.046 335.3 119.092C336.811 120.008 338.334 120.466 339.868 120.466C341.082 120.466 342.032 120.203 342.719 119.676C343.428 119.127 343.783 118.36 343.783 117.375C343.783 116.78 343.623 116.299 343.302 115.933C343.005 115.543 342.604 115.234 342.1 115.005C341.597 114.753 340.807 114.421 339.731 114.009C338.334 113.483 337.235 112.99 336.434 112.533C335.632 112.075 335.003 111.491 334.545 110.781C334.087 110.071 333.858 109.167 333.858 108.068C333.858 106.305 334.442 104.965 335.609 104.05C336.777 103.134 338.277 102.676 340.108 102.676C341.07 102.676 342.02 102.813 342.959 103.088C343.921 103.363 344.779 103.752 345.535 104.256C345.626 104.301 345.684 104.382 345.707 104.496C345.729 104.611 345.718 104.714 345.672 104.805L344.985 106.213C344.94 106.351 344.848 106.442 344.711 106.488C344.596 106.534 344.47 106.511 344.333 106.419C342.867 105.549 341.459 105.114 340.108 105.114C338.987 105.114 338.105 105.355 337.464 105.835C336.846 106.316 336.537 107.015 336.537 107.93C336.537 108.64 336.708 109.213 337.052 109.648C337.395 110.083 337.807 110.415 338.288 110.644C338.769 110.85 339.536 111.147 340.589 111.537C341.94 112.017 343.005 112.464 343.783 112.876C344.562 113.288 345.203 113.838 345.707 114.525C346.21 115.211 346.462 116.104 346.462 117.203C346.462 118.966 345.844 120.363 344.608 121.393C343.394 122.401 341.768 122.904 339.731 122.904ZM367.043 123.008C365.692 123.008 364.444 122.676 363.3 122.012C362.155 121.348 361.25 120.443 360.586 119.298C359.922 118.131 359.59 116.871 359.59 115.52V110.163C359.59 108.812 359.922 107.564 360.586 106.419C361.25 105.252 362.155 104.336 363.3 103.672C364.444 103.008 365.692 102.676 367.043 102.676C368.394 102.676 369.642 103.008 370.787 103.672C371.931 104.336 372.836 105.252 373.5 106.419C374.164 107.564 374.496 108.812 374.496 110.163V115.52C374.496 116.871 374.164 118.131 373.5 119.298C372.836 120.443 371.931 121.348 370.787 122.012C369.642 122.676 368.394 123.008 367.043 123.008ZM367.043 120.466C367.913 120.466 368.714 120.26 369.447 119.848C370.18 119.413 370.752 118.829 371.164 118.096C371.599 117.364 371.817 116.562 371.817 115.692V109.991C371.817 109.121 371.599 108.32 371.164 107.587C370.752 106.854 370.18 106.282 369.447 105.87C368.714 105.435 367.913 105.217 367.043 105.217C366.173 105.217 365.372 105.435 364.639 105.87C363.906 106.282 363.322 106.854 362.887 107.587C362.475 108.32 362.269 109.121 362.269 109.991V115.692C362.269 116.562 362.475 117.364 362.887 118.096C363.322 118.829 363.906 119.413 364.639 119.848C365.372 120.26 366.173 120.466 367.043 120.466ZM380.315 122.527C380.178 122.527 380.064 122.481 379.972 122.389C379.88 122.298 379.835 122.183 379.835 122.046V103.603C379.835 103.466 379.88 103.351 379.972 103.26C380.064 103.168 380.178 103.122 380.315 103.122H382.033C382.17 103.122 382.285 103.168 382.376 103.26C382.468 103.351 382.513 103.466 382.513 103.603V105.492C383.109 104.553 383.876 103.843 384.815 103.363C385.776 102.882 386.852 102.641 388.043 102.641C389.989 102.641 391.557 103.237 392.748 104.427C393.962 105.618 394.568 107.186 394.568 109.132V122.046C394.568 122.183 394.522 122.298 394.431 122.389C394.339 122.481 394.225 122.527 394.087 122.527H392.37C392.233 122.527 392.118 122.481 392.027 122.389C391.935 122.298 391.889 122.183 391.889 122.046V109.648C391.889 108.182 391.5 107.072 390.722 106.316C389.966 105.538 388.936 105.149 387.631 105.149C386.097 105.149 384.86 105.652 383.922 106.66C382.983 107.667 382.513 109.052 382.513 110.815V122.046C382.513 122.183 382.468 122.298 382.376 122.389C382.285 122.481 382.17 122.527 382.033 122.527H380.315ZM410.72 122.561C410.376 122.561 410.079 122.447 409.827 122.218C409.598 121.966 409.483 121.668 409.483 121.325V99.9283C409.483 99.5848 409.598 99.2986 409.827 99.0697C410.079 98.8178 410.376 98.6919 410.72 98.6919H416.833C418.459 98.6919 419.958 99.104 421.332 99.9283C422.706 100.73 423.793 101.829 424.595 103.225C425.396 104.599 425.797 106.11 425.797 107.759V113.494C425.797 115.143 425.396 116.665 424.595 118.062C423.793 119.436 422.706 120.535 421.332 121.359C419.958 122.16 418.459 122.561 416.833 122.561H410.72ZM416.833 119.917C417.955 119.917 418.985 119.642 419.924 119.092C420.863 118.543 421.607 117.799 422.156 116.86C422.706 115.898 422.981 114.857 422.981 113.735V107.484C422.981 106.362 422.706 105.332 422.156 104.393C421.607 103.454 420.863 102.71 419.924 102.161C418.985 101.611 417.955 101.336 416.833 101.336H412.3V119.917H416.833ZM437.71 122.561C436.359 122.561 435.111 122.229 433.967 121.565C432.822 120.901 431.917 119.997 431.253 118.852C430.589 117.684 430.257 116.425 430.257 115.074V110.163C430.257 108.812 430.589 107.564 431.253 106.419C431.917 105.252 432.822 104.336 433.967 103.672C435.111 103.008 436.359 102.676 437.71 102.676C439.061 102.676 440.309 103.008 441.454 103.672C442.598 104.336 443.503 105.252 444.167 106.419C444.831 107.564 445.163 108.812 445.163 110.163V112.498C445.163 112.842 445.037 113.139 444.785 113.391C444.556 113.62 444.27 113.735 443.926 113.735H432.936V115.246C432.936 116.116 433.142 116.917 433.555 117.65C433.99 118.382 434.573 118.966 435.306 119.401C436.039 119.814 436.84 120.02 437.71 120.02H442.69C442.827 120.02 442.942 120.065 443.033 120.157C443.125 120.249 443.171 120.363 443.171 120.5V122.08C443.171 122.218 443.125 122.332 443.033 122.424C442.942 122.515 442.827 122.561 442.69 122.561H437.71ZM442.484 111.262V109.991C442.484 109.121 442.266 108.32 441.831 107.587C441.419 106.854 440.847 106.282 440.114 105.87C439.382 105.435 438.58 105.217 437.71 105.217C436.84 105.217 436.039 105.435 435.306 105.87C434.573 106.282 433.99 106.854 433.555 107.587C433.142 108.32 432.936 109.121 432.936 109.991V111.262H442.484ZM450.714 122.561C450.577 122.561 450.462 122.515 450.371 122.424C450.279 122.332 450.233 122.218 450.233 122.08V103.603C450.233 103.466 450.279 103.351 450.371 103.26C450.462 103.168 450.577 103.122 450.714 103.122H452.431C452.569 103.122 452.683 103.168 452.775 103.26C452.866 103.351 452.912 103.466 452.912 103.603V105.458C453.462 104.542 454.183 103.855 455.076 103.397C455.992 102.916 456.999 102.676 458.098 102.676C459.541 102.676 460.788 103.019 461.842 103.706C462.918 104.37 463.685 105.286 464.143 106.454C464.624 105.24 465.368 104.313 466.375 103.672C467.383 103.031 468.527 102.71 469.81 102.71C471.756 102.71 473.324 103.305 474.515 104.496C475.728 105.687 476.335 107.255 476.335 109.201V122.08C476.335 122.218 476.289 122.332 476.198 122.424C476.106 122.515 475.991 122.561 475.854 122.561H474.137C474 122.561 473.885 122.515 473.793 122.424C473.702 122.332 473.656 122.218 473.656 122.08V109.716C473.656 108.251 473.267 107.141 472.488 106.385C471.733 105.606 470.702 105.217 469.397 105.217C467.978 105.217 466.822 105.721 465.929 106.728C465.059 107.736 464.624 109.121 464.624 110.884V122.08C464.624 122.218 464.578 122.332 464.486 122.424C464.395 122.515 464.28 122.561 464.143 122.561H462.426C462.288 122.561 462.174 122.515 462.082 122.424C461.991 122.332 461.945 122.218 461.945 122.08V109.682C461.945 108.217 461.556 107.106 460.777 106.351C460.021 105.572 458.991 105.183 457.686 105.183C456.267 105.183 455.11 105.687 454.217 106.694C453.347 107.701 452.912 109.087 452.912 110.85V122.08C452.912 122.218 452.866 122.332 452.775 122.424C452.683 122.515 452.569 122.561 452.431 122.561H450.714ZM486.631 123.008C484.868 123.008 483.426 122.515 482.304 121.531C481.182 120.523 480.621 119.207 480.621 117.581C480.621 116.07 481.125 114.822 482.132 113.838C483.163 112.853 484.594 112.235 486.425 111.983L490.89 111.365C491.554 111.273 492.035 111.056 492.333 110.712C492.63 110.346 492.779 109.831 492.779 109.167V108.892C492.779 107.793 492.344 106.912 491.474 106.248C490.627 105.561 489.562 105.217 488.28 105.217C487.273 105.217 486.391 105.423 485.636 105.835C484.903 106.225 484.319 106.774 483.884 107.484C483.701 107.782 483.472 107.93 483.197 107.93C483.06 107.93 482.911 107.873 482.751 107.759L481.823 107.141C481.617 106.957 481.514 106.763 481.514 106.557C481.514 106.465 481.549 106.351 481.617 106.213C482.281 105.114 483.186 104.256 484.33 103.637C485.498 102.996 486.838 102.676 488.349 102.676C489.722 102.676 490.947 102.939 492.024 103.466C493.1 103.969 493.924 104.691 494.496 105.629C495.092 106.568 495.389 107.656 495.389 108.892V122.08C495.389 122.218 495.343 122.332 495.252 122.424C495.16 122.515 495.046 122.561 494.908 122.561H493.466C493.329 122.561 493.214 122.515 493.123 122.424C493.031 122.332 492.985 122.218 492.985 122.08V119.71C492.275 120.741 491.36 121.554 490.238 122.149C489.139 122.721 487.937 123.008 486.631 123.008ZM486.906 120.603C488.028 120.603 489.047 120.306 489.963 119.71C490.879 119.115 491.588 118.371 492.092 117.478C492.619 116.562 492.882 115.669 492.882 114.799V113.391C492.333 113.46 490.57 113.7 487.593 114.112L486.631 114.25C484.342 114.593 483.197 115.669 483.197 117.478C483.197 118.417 483.541 119.172 484.227 119.745C484.914 120.317 485.807 120.603 486.906 120.603ZM501.694 122.527C501.556 122.527 501.442 122.481 501.35 122.389C501.259 122.298 501.213 122.183 501.213 122.046V103.603C501.213 103.466 501.259 103.351 501.35 103.26C501.442 103.168 501.556 103.122 501.694 103.122H503.411C503.548 103.122 503.663 103.168 503.754 103.26C503.846 103.351 503.892 103.466 503.892 103.603V105.492C504.487 104.553 505.254 103.843 506.193 103.363C507.155 102.882 508.231 102.641 509.421 102.641C511.367 102.641 512.936 103.237 514.126 104.427C515.34 105.618 515.947 107.186 515.947 109.132V122.046C515.947 122.183 515.901 122.298 515.809 122.389C515.718 122.481 515.603 122.527 515.466 122.527H513.749C513.611 122.527 513.497 122.481 513.405 122.389C513.314 122.298 513.268 122.183 513.268 122.046V109.648C513.268 108.182 512.879 107.072 512.1 106.316C511.344 105.538 510.314 105.149 509.009 105.149C507.475 105.149 506.239 105.652 505.3 106.66C504.361 107.667 503.892 109.052 503.892 110.815V122.046C503.892 122.183 503.846 122.298 503.754 122.389C503.663 122.481 503.548 122.527 503.411 122.527H501.694ZM528.092 123.008C526.787 123.008 525.585 122.687 524.486 122.046C523.387 121.405 522.517 120.535 521.876 119.436C521.234 118.337 520.914 117.135 520.914 115.83V109.854C520.914 108.549 521.234 107.347 521.876 106.248C522.517 105.149 523.387 104.279 524.486 103.637C525.608 102.996 526.821 102.676 528.126 102.676C529.065 102.676 529.969 102.848 530.839 103.191C531.732 103.534 532.511 104.015 533.175 104.633V96.9747C533.175 96.8373 533.221 96.7228 533.312 96.6312C533.404 96.5396 533.518 96.4938 533.656 96.4938H535.373C535.51 96.4938 535.625 96.5396 535.716 96.6312C535.808 96.7228 535.854 96.8373 535.854 96.9747V122.08C535.854 122.218 535.808 122.332 535.716 122.424C535.625 122.515 535.51 122.561 535.373 122.561H533.793C533.656 122.561 533.541 122.515 533.449 122.424C533.358 122.332 533.312 122.218 533.312 122.08V120.569C532.671 121.348 531.893 121.954 530.977 122.389C530.084 122.801 529.122 123.008 528.092 123.008ZM528.367 120.466C529.237 120.466 530.038 120.26 530.771 119.848C531.526 119.413 532.11 118.829 532.522 118.096C532.957 117.364 533.175 116.562 533.175 115.692V109.991C533.175 109.121 532.957 108.32 532.522 107.587C532.11 106.854 531.526 106.282 530.771 105.87C530.038 105.435 529.237 105.217 528.367 105.217C527.497 105.217 526.695 105.435 525.962 105.87C525.23 106.282 524.646 106.854 524.211 107.587C523.799 108.32 523.593 109.121 523.593 109.991V115.692C523.593 116.562 523.799 117.364 524.211 118.096C524.646 118.829 525.23 119.413 525.962 119.848C526.695 120.26 527.497 120.466 528.367 120.466Z" fill="#F1F3F4"/>
<defs>
<filter id="filter0_d_2027_657" x="178.086" y="98.2338" width="393.829" height="58.1472" filterUnits="userSpaceOnUse" color-interpolation-filters="sRGB">
<feFlood flood-opacity="0" result="BackgroundImageFix"/>
<feColorMatrix in="SourceAlpha" type="matrix" values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 127 0" result="hardAlpha"/>
<feMorphology radius="2.17239" operator="dilate" in="SourceAlpha" result="effect1_dropShadow_2027_657"/>
<feOffset dy="4.34478"/>
<feGaussianBlur stdDeviation="10.6447"/>
<feComposite in2="hardAlpha" operator="out"/>
<feColorMatrix type="matrix" values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0.15 0"/>
<feBlend mode="normal" in2="BackgroundImageFix" result="effect1_dropShadow_2027_657"/>
<feBlend mode="normal" in="SourceGraphic" in2="effect1_dropShadow_2027_657" result="shape"/>
</filter>
<linearGradient id="paint0_linear_2027_657" x1="238.281" y1="67.6771" x2="385.532" y2="221.101" gradientUnits="userSpaceOnUse">
<stop stop-color="#4E3939"/>
<stop offset="1" stop-color="white" stop-opacity="0.27"/>
</linearGradient>
<clipPath id="clip0_2027_657">
<rect width="139" height="37" fill="white" transform="translate(39.5693 34.2729)"/>
</clipPath>
</defs>
</svg>
'@

    $svgRJLogo_dark = @'
<?xml version="1.0" encoding="utf-8"?>
<!-- Generator: Adobe Illustrator 27.6.0, SVG Export Plug-In . SVG Version: 6.00 Build 0)  -->
<svg version="1.1" id="Realmjoin" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px"
	 viewBox="0 0 174.6 47" style="enable-background:new 0 0 174.6 47;" xml:space="preserve">
<style type="text/css">
	.st0{display:none;opacity:0.78;fill:#595959;}
	.st1{fill:#FFFFFF;}
	.st2{fill:#F8842C;}
</style>
<rect class="st0" width="174.6" height="47"/>
<path class="st1" d="M48.8,42.5c-0.1-0.1-0.1-0.2-0.1-0.3V23.5c0-0.3,0.1-0.5,0.3-0.7s0.5-0.3,0.7-0.3h5c2.1,0,4.1,1.1,5.1,2.9
	c0.5,0.9,0.8,1.9,0.8,2.9c0,1.4-0.4,2.7-1.3,3.7c-0.9,1-2,1.7-3.4,2l4.8,7.8c0.1,0.1,0.1,0.2,0.1,0.3s-0.1,0.3-0.2,0.4
	c-0.1,0.1-0.2,0.1-0.4,0.1h-1.5c-0.2,0-0.4-0.1-0.4-0.3l-5.1-8.2H51v8.1c0,0.1,0,0.2-0.1,0.3s-0.2,0.1-0.3,0.1h-1.5
	C49,42.6,48.8,42.5,48.8,42.5z M54.6,31.9c1,0,1.9-0.3,2.6-1c0.7-0.7,1-1.6,1-2.5c0-1-0.3-1.9-1-2.6s-1.6-1-2.6-1H51V32L54.6,31.9z"
	/>
<path class="st1" d="M65.9,41.7c-1-0.5-1.7-1.3-2.3-2.3c-0.6-1-0.8-2.1-0.8-3.2V32c0-1.1,0.3-2.2,0.8-3.2c1.1-2,3.2-3.2,5.4-3.1
	c1.1,0,2.2,0.3,3.2,0.8s1.7,1.3,2.3,2.3c0.6,1,0.8,2.1,0.8,3.2v2c0,0.3-0.1,0.5-0.3,0.7c-0.2,0.2-0.5,0.3-0.7,0.3H65v1.2
	c0,0.7,0.2,1.4,0.5,2s0.9,1.1,1.5,1.5c0.6,0.3,1.3,0.5,2,0.5h4.2c0.1,0,0.2,0,0.3,0.1c0.1,0.1,0.1,0.2,0.1,0.3v1.3
	c0,0.2-0.2,0.4-0.4,0.4l0,0H69C67.9,42.6,66.8,42.3,65.9,41.7z M73.1,33v-1c0-0.7-0.2-1.4-0.5-2s-0.8-1.1-1.5-1.5
	c-1.9-1.1-4.4-0.5-5.5,1.4V30c-0.4,0.6-0.6,1.3-0.6,2v1H73.1z"/>
<path class="st1" d="M78.5,41.7c-1.8-1.7-1.9-4.6-0.2-6.5l0,0c1-0.9,2.3-1.4,3.6-1.6l3.8-0.5c0.5,0,0.9-0.2,1.2-0.6
	c0.3-0.4,0.4-0.8,0.4-1.3V31c0-0.9-0.4-1.7-1.1-2.2c-0.8-0.6-1.7-0.9-2.7-0.9c-0.8,0-1.5,0.2-2.2,0.5c-0.6,0.3-1.1,0.8-1.5,1.4
	c-0.1,0.2-0.3,0.4-0.6,0.4c-0.1,0-0.3-0.1-0.4-0.1L78,29.6c-0.2-0.1-0.3-0.3-0.3-0.5c0-0.1,0-0.2,0.1-0.3c0.6-0.9,1.4-1.7,2.3-2.2
	c1-0.6,2.2-0.8,3.4-0.8c1.1,0,2.1,0.2,3.1,0.6c0.9,0.4,1.6,1,2.1,1.8s0.8,1.8,0.7,2.8v11.2c0,0.2-0.2,0.4-0.4,0.4l0,0h-1.2
	c-0.1,0-0.2,0-0.3-0.1c-0.1-0.1-0.1-0.2-0.1-0.3v-2c-0.6,0.8-1.4,1.5-2.3,2s-2,0.7-3.1,0.7C80.8,43,79.5,42.5,78.5,41.7z M84.9,40.2
	c0.8-0.5,1.4-1.1,1.8-1.9c0.4-0.7,0.6-1.5,0.7-2.3v-1.2l-4.5,0.6l-0.8,0.1c-1.9,0.3-2.9,1.2-2.9,2.7c0,0.7,0.3,1.4,0.9,1.9
	s1.4,0.8,2.3,0.7C83.3,40.9,84.2,40.7,84.9,40.2z"/>
<path class="st1" d="M94.5,42.3c-0.2-0.2-0.3-0.5-0.3-0.7V22.7h-2.7c-0.1,0-0.2,0-0.3-0.1c-0.1-0.1-0.1-0.2-0.1-0.3V21
	c0-0.1,0-0.2,0.1-0.3c0.1-0.1,0.2-0.1,0.3-0.1h3.9c0.3,0,0.5,0.1,0.7,0.3c0.2,0.2,0.3,0.5,0.3,0.7v18.8h3.1c0.1,0,0.2,0,0.3,0.1
	c0.1,0.1,0.1,0.2,0.1,0.3v1.3c0,0.2-0.2,0.4-0.4,0.4l0,0h-4.3C95,42.6,94.7,42.5,94.5,42.3z"/>
<path class="st1" d="M102.4,42.5c-0.1-0.1-0.1-0.2-0.1-0.3V26.6c0-0.1,0-0.2,0.1-0.3c0.1-0.1,0.2-0.1,0.3-0.1h1.4
	c0.1,0,0.2,0,0.3,0.1c0.1,0.1,0.1,0.2,0.1,0.3v1.6c0.4-0.7,1.1-1.4,1.8-1.8c0.8-0.4,1.7-0.6,2.5-0.6c1.1,0,2.2,0.3,3.2,0.9
	c0.9,0.6,1.5,1.4,1.9,2.4c0.4-1,1-1.8,1.9-2.4c0.9-0.5,1.9-0.8,2.9-0.8c1.5-0.1,2.9,0.5,4,1.5c1,1.1,1.6,2.5,1.5,4v10.8
	c0,0.2-0.2,0.4-0.4,0.4l0,0h-1.4c-0.2,0-0.4-0.2-0.4-0.4l0,0V31.7c0.1-1-0.3-2.1-1-2.8c-0.7-0.7-1.7-1-2.6-1c-1.1,0-2.2,0.4-2.9,1.3
	c-0.8,1-1.2,2.2-1.1,3.5v9.5c0,0.2-0.2,0.4-0.4,0.4l0,0h-1.4c-0.2,0-0.4-0.2-0.4-0.4l0,0V31.7c0.1-1-0.3-2.1-1-2.8
	c-0.7-0.7-1.6-1-2.6-1c-1.1,0-2.2,0.4-2.9,1.3c-0.8,1-1.2,2.2-1.1,3.5v9.5c0,0.1,0,0.2-0.1,0.3s-0.2,0.1-0.3,0.1h-1.4
	C102.5,42.5,102.5,42.5,102.4,42.5z"/>
<path class="st1" d="M126.3,42.5c-0.1-0.1-0.1-0.2-0.1-0.3v-1.3c0-0.2,0.1-0.4,0.3-0.4h0.1h1.7c0.6,0,1.2-0.2,1.6-0.6
	c0.4-0.5,0.7-1.1,0.6-1.7V22.8c0-0.1,0-0.2,0.1-0.3c0.1-0.1,0.2-0.1,0.3-0.1h1.6c0.1,0,0.2,0,0.3,0.1s0.1,0.2,0.1,0.3V38
	c0,0.8-0.2,1.6-0.5,2.3s-0.8,1.3-1.5,1.7s-1.5,0.6-2.3,0.6h-2C126.5,42.6,126.4,42.6,126.3,42.5z"/>
<path class="st1" d="M139,42.1c-1-0.6-1.7-1.3-2.3-2.3s-0.8-2.1-0.8-3.2v-4.5c-0.1-3.5,2.6-6.4,6-6.5c3.5-0.1,6.4,2.6,6.5,6
	c0,0.2,0,0.3,0,0.5v4.5c0,1.1-0.3,2.2-0.8,3.2C145.9,42.8,142,43.9,139,42.1z M144.2,40.3c0.6-0.3,1.1-0.9,1.5-1.5s0.5-1.3,0.5-2V32
	c0-0.7-0.2-1.4-0.5-2s-0.9-1.1-1.5-1.5c-0.6-0.3-1.3-0.5-2-0.5s-1.4,0.2-2,0.5c-0.6,0.4-1.1,0.9-1.5,1.5s-0.5,1.3-0.5,2v4.8
	c0,0.7,0.2,1.4,0.5,2s0.9,1.1,1.5,1.5s1.3,0.5,2,0.5S143.6,40.6,144.2,40.3L144.2,40.3z"/>
<path class="st1" d="M153.4,42.3c-0.2-0.2-0.3-0.5-0.3-0.7V28.3h-3c-0.2,0-0.4-0.1-0.4-0.4l0,0v-1.3c0-0.2,0.1-0.4,0.4-0.4l0,0h4.2
	c0.3,0,0.5,0.1,0.7,0.3c0.2,0.2,0.3,0.5,0.3,0.7v13.2h3c0.2,0,0.4,0.2,0.4,0.4l0,0v1.3c0,0.1,0,0.2-0.1,0.3s-0.2,0.1-0.3,0.1h-4.2
	C153.8,42.6,153.6,42.5,153.4,42.3z M153.2,23.8c-0.1-0.1-0.1-0.2-0.1-0.3V21c0-0.1,0-0.2,0.1-0.3c0.1-0.1,0.2-0.1,0.3-0.1h1.4
	c0.1,0,0.2,0,0.3,0.1s0.1,0.2,0.1,0.3v2.5c0,0.1,0,0.2-0.1,0.3c-0.1,0.1-0.2,0.1-0.3,0.1h-1.4C153.4,23.9,153.3,23.9,153.2,23.8
	L153.2,23.8z"/>
<path class="st1" d="M161.3,42.4c-0.1-0.1-0.1-0.2-0.1-0.3V26.6c0-0.2,0.2-0.4,0.4-0.4l0,0h1.4c0.2,0,0.4,0.1,0.4,0.4l0,0v1.6
	c0.5-0.8,1.2-1.4,2-1.8s1.8-0.6,2.7-0.6c1.5-0.1,2.9,0.5,4,1.5c1,1.1,1.6,2.5,1.5,4v10.9c0,0.2-0.1,0.4-0.4,0.4l0,0h-1.4
	c-0.2,0-0.4-0.2-0.4-0.4l0,0V31.7c0.1-1-0.3-2.1-1-2.8c-0.7-0.7-1.7-1-2.6-1c-1.2,0-2.3,0.4-3.1,1.3c-0.8,1-1.3,2.2-1.2,3.5v9.5
	c0,0.2-0.1,0.4-0.4,0.4l0,0h-1.4C161.5,42.5,161.4,42.5,161.3,42.4z"/>
<g>
	<path class="st2" d="M22.2,39.5c0-0.4,0.3-0.8,0.7-0.8c0,0,0,0,0.1,0h2.8c0.6,0,1.1-0.2,1.5-0.6c0.4-0.4,0.6-1,0.5-1.6V16.7
		c0-0.4,0.3-0.8,0.8-0.8l0,0h3c0.4,0,0.8,0.3,0.8,0.7c0,0,0,0,0,0.1v19.8c0,1.2-0.3,2.3-0.8,3.4c-0.5,1-1.3,1.8-2.3,2.3
		c-1,0.6-2.1,0.9-3.3,0.9h-2.2L40,46l-4.8-31L5.8,1L3.9,16.2c0.3-0.2,0.6-0.3,0.9-0.3l0,0h7.1c1.5,0,3,0.3,4.2,1.1
		c1.3,0.7,2.3,1.7,3.1,3c1.7,2.8,1.5,6.4-0.5,9c-1.1,1.4-2.6,2.5-4.3,3l6.3,9.8c0.1,0.2,0.1,0.3,0.2,0.5c0,0.1,0,0.2-0.1,0.3
		l1.5,0.3c-0.1-0.1-0.2-0.3-0.2-0.5v-2.9H22.2z"/>
	<path class="st1" d="M26.1,43.1c1.1,0,2.3-0.3,3.3-0.9c1-0.6,1.7-1.4,2.3-2.3c0.5-1,0.8-2.2,0.8-3.4V16.7c0,0,0,0,0-0.1
		c0-0.4-0.4-0.8-0.8-0.7h-3l0,0c-0.4,0-0.8,0.4-0.8,0.8v19.8c0,0.6-0.2,1.2-0.5,1.6c-0.4,0.4-0.9,0.6-1.5,0.6H23c0,0,0,0-0.1,0
		c-0.4,0-0.8,0.4-0.7,0.8v2.8c0,0.2,0.1,0.4,0.2,0.5l1.5,0.3C23.9,43.1,26.1,43.1,26.1,43.1z"/>
	<polygon class="st2" points="7.8,32.4 7.8,40.2 15.4,41.6 9.7,32.4 	"/>
	<polygon class="st2" points="1,39 3.2,39.4 3.2,21.6 	"/>
	<path class="st2" d="M15.7,24.2c0-1.1-0.4-2.1-1.1-2.8c-0.8-0.7-1.8-1.1-2.9-1.1H7.8v7.8h3.9c1.1,0,2.1-0.3,2.9-1.1
		C15.3,26.3,15.7,25.3,15.7,24.2z"/>
	<path class="st1" d="M21,42.3c0-0.2-0.1-0.4-0.2-0.5L14.5,32c1.7-0.5,3.2-1.6,4.3-3c2-2.6,2.2-6.2,0.5-9c-0.7-1.3-1.8-2.3-3.1-3
		s-2.8-1.1-4.2-1.1H4.8l0,0c-0.3,0-0.7,0.1-0.9,0.3l-0.7,5.4v17.8l4.6,0.8v-7.8h1.9l5.7,9.2l5.5,1C20.9,42.5,21,42.4,21,42.3z
		 M7.8,28.1v-7.8h3.9c1.1,0,2.1,0.4,2.9,1.1c0.7,0.7,1.2,1.8,1.1,2.8c0,1-0.4,2-1.1,2.8c-0.8,0.7-1.8,1.1-2.9,1.1L7.8,28.1z"/>
</g>
</svg>
'@

    $plainBase64Header = [Convert]::ToBase64String(
        [Text.Encoding]::UTF8.GetBytes($svgRJHeader)
    )

    $plainBase64Logo_dark = [Convert]::ToBase64String(
        [Text.Encoding]::UTF8.GetBytes($svgRJLogo_dark)
    )

    return @"
<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="light">
    <meta name="supported-color-schemes" content="light">
    <title>$Subject</title>
    <!--[if !mso]><!-->
    <link href="https://fonts.googleapis.com/css2?family=Miriam+Libre:wght@400;700&display=swap" rel="stylesheet">
    <!--<![endif]-->
    <!-- Base styles for ALL clients (including Dark Mode for modern clients) -->
<style type="text/css">
    /* === RESET & BASICS === */
    * { margin: 0; padding: 0; box-sizing: border-box; }

    body {
        font-family: 'Miriam Libre', -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
        line-height: 1.6;
        color: #011e33;
        background-color: #e8ebed;
        padding: 20px;
    }

    /* === CONTAINER === */
    .email-container {
        max-width: 750px;
        margin: 0 auto;
        background-color: #ffffff;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        overflow: hidden;
    }

    /* === HEADER === */
    .header {
        background: url('data:image/svg+xml;base64,$($plainBase64Header)') no-repeat center top;
        background-size: contain;
        width: 100%;
        padding-top: 26.67%;
        display: block;
        position: relative;
    }

    /* === CONTENT === */
    .content {
        padding: 48px;
        background-color: #ffffff;
    }

    .tenant-info {
        background: #e8ebed;
        border: 1px solid #e0e7ff;
        border-left: 4px solid #f8842c;
        padding: 10px 20px;
        margin-top: 32px;
        border-radius: 8px;
        font-size: 14px;
    }

    .tenant-info strong {
        color: #011e33;
        font-weight: 600;
    }

    .content h1 {
        color: #111827;
        border-bottom: 2px solid #111827;
        padding-bottom: 12px;
        margin-bottom: 15px;
        font-size: 26px;
        line-height: 1.4;
        font-weight: 800;
    }

    .content h2 {
        color: #111827;
        margin-top: 42px;
        margin-bottom: 15px;
        font-size: 22px;
        line-height: 1.4;
        font-weight: 800;
    }

    .content h3 {
        color: #111827;
        margin-top: 27px;
        margin-bottom: 15px;
        font-size: 18px;
        line-height: 1.4;
        font-weight: 800;
    }

    .content h4,
    .content h5 {
        color: #111827;
        margin-top: 15px;
        margin-bottom: 15px;
        font-size: 16px;
        line-height: 1.4;
        font-weight: 800;
    }

    .content p {
        color: #111827;
        font-size: 16px;
        line-height: 1.4;
        margin-bottom: 15px;
    }

    .content ul {
        margin-top: 15px;
        margin-left: 0;
        margin-bottom: 12px;
        list-style-type: disc;
        padding-left: 20px;
    }

    .content ol {
        margin-top: 15px;
        margin-left: 0;
        margin-bottom: 12px;
        list-style-type: decimal;
        padding-left: 20px;
    }

    .content li {
        margin-top: 4px;
        margin-left: 0;
        color: #011e33;
        line-height: 1.5;
        margin-bottom: 4px;
        padding-left: 8px;
    }

    /* === TABLES === */
    .table-wrapper {
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
        margin: 15px 0;
    }

    .content table {
        width: 100%;
        border-collapse: collapse;
        margin: 0;
        background-color: white;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);
    }

    .content th {
        background: #f8842c !important;
        color: #ffffff !important;
        padding: 8px 16px;
        text-align: left;
        font-weight: 600;
        font-size: 14px;
        text-transform: uppercase;
    }

    .content td {
        padding: 8px 16px;
        border-bottom: 1px solid #e8ebed;
        font-size: 14px;
        color: #2D3748;
    }

    .content tr:nth-child(even) {
        background-color: #e8ebed;
    }

    .content blockquote {
        border-left: 4px solid #3b82f6;
        background: #e8ebed;
        padding: 20px 24px;
        margin-top: 15px;
        margin-bottom: 15px;
        border-radius: 0 8px 8px 0;
        font-style: italic;
        color: #374151;
    }

    /* Fix spacing after blockquote */
    .content blockquote + p {
        margin-top: 15px !important;
    }

    /* === CODE === */
    .content code {
        background-color: #e8ebed;
        padding: 2px 8px;
        border-radius: 4px;
        font-family: 'SF Mono', Monaco, 'Consolas', monospace;
        font-size: 0.875em;
        color: #011e33;
        border: 1px solid #e5e7eb;
    }

    .content pre {
        background-color: #e8ebed;
        padding: 20px;
        border-radius: 8px;
        overflow-x: auto;
        margin: 20px 0;
        border: 1px solid #e5e7eb;
        font-family: 'SF Mono', Monaco, 'Consolas', monospace;
    }

    /* === HR (Horizontal Rule) === */
    .content hr {
        border: none;
        border-top: 2px solid #e5e7eb;
        margin: 24px 0;
    }

    /* === STRONG & EM === */
    .content strong {
        font-weight: 600;
        color: #111827;
    }

    .content em {
        font-style: italic;
        color: #374151;
    }

    /* === LINKS === */
    .content a {
        color: #3b82f6;
        text-decoration: underline;
    }

    .content a:hover {
        color: #2563eb;
    }

    /* === ATTACHMENTS === */
    .attachments {
        background: #e8ebed;
        border: 1px solid #e0e7ff;
        border-left: 4px solid #f8842c;
        border-radius: 8px;
        padding: 10px 20px;
        margin-top: 10px;
    }

    .attachments h3 {
        color: #011e33;
        margin-top: 0;
        font-size: 14px;
        font-weight: 600;
    }

    .attachment-list {
        margin: 0 0 16px 0;
        padding: 0;
    }

    .attachment-item {
        background-color: white;
        border: 1px solid #e0e7ff;
        border-radius: 6px;
        padding: 8px 12px;
        margin-bottom: 3px;
        font-size: 14px;
    }

    .attachments p {
        margin-bottom: 0;
        font-size: 14px;
    }

    /* === FOOTER === */
    .footer {
        background: #011e33;
        color: #3f3f3f;
        padding: 40px 48px;
        text-align: center;
    }

    .footer .logo-container {
        margin-bottom: 16px;
        max-width: 130px;
        margin-left: auto;
        margin-right: auto;
    }

    .footer .logo-dark {
        max-width: 130px;
        width: 130px !important;
        height: auto;
        opacity: 0.9;
        display: block;
        margin: 0 auto;
    }

    .footer .tagline {
        font-size: 14px;
        opacity: 0.8;
        margin-bottom: 10px;
        color: #ffffff;
    }

    .footer .links {
        font-size: 13px;
        opacity: 0.7;
    }

    .footer .links a {
        color: #60a5fa;
        text-decoration: none;
        margin: 0 12px;
    }

        @media (max-width: 768px) {
        body { padding: 10px; }
        .email-container {
            max-width: 100%;
            border-radius: 8px;
        }
        .content, .footer { padding: 24px 20px; }
        .footer .logo-dark { max-width: 120px !important; }
        .content h1 { font-size: 22px; line-height: 1.4; }
        .content h2 { font-size: 18px; line-height: 1.4; }
        .content h3 { font-size: 16px; line-height: 1.4; }
        .content h4, .content h5 { font-size: 16px; line-height: 1.4; }
        .content p { font-size: 16px; line-height: 1.4; }
        .table-wrapper { margin: 15px 0; }
        .content table { font-size: 13px; min-width: 500px; }
        .content th, .content td { padding: 6px 8px; }
        .tenant-info, .attachments { padding: 16px 20px; font-size: 13px; }
    }

    /* === TABLET === */
    @media (min-width: 769px) and (max-width: 1024px) {
        .email-container { max-width: 750px; }
        .content, .footer { padding: 36px; }
        .footer .logo-dark { max-width: 160px; }
    }

    /* === DESKTOP === */
    @media (min-width: 1025px) {
        .email-container { max-width: 750px; }
        .footer .logo-dark { max-width: 160px; }
    }

    /* === DARK MODE (New Outlook, modern clients) === */
    @media (prefers-color-scheme: dark) {
        body { background-color: #1a1a1a !important; }

        .email-container, .content {
            background-color: #2d2d2d !important;
            color: #e5e5e5 !important;
        }

        .header {
            /* Keep Light Mode header graphic in Dark Mode */
            background: url('data:image/svg+xml;base64,$($plainBase64Header)') no-repeat center center !important;
            background-size: cover !important;
        }

        h1, h2, h3, p, span, strong, div, li, td, blockquote {
            color: #e5e5e5 !important;
        }

        h1 {
            border-bottom: 2px solid #e5e5e5 !important;
        }

        .tenant-info {
            background: linear-gradient(135deg, #2d2d2d 0%, #3a3a3a 100%) !important;
            border: 1px solid #4a4a4a !important;
            border-left-color: #f8842c !important;
        }

        .content table {
            background-color: #3a3a3a !important;
        }

        .content td {
            border-bottom-color: #4a4a4a !important;
        }

        .content th {
            background: #f8842c !important;
            color: #ffffff !important;
        }

        .content tr:nth-child(even) {
            background-color: #404040 !important;
        }

        .attachments {
            background: linear-gradient(135deg, #2d2d2d 0%, #3a3a3a 100%) !important;
            border: 1px solid #4a4a4a !important;
            border-left-color: #f8842c !important;
        }

        .attachment-item {
            background-color: #2d2d2d !important;
            border-color: #4a4a4a !important;
        }

        .content code, .content pre {
            background-color: #404040 !important;
            color: #e5e5e5 !important;
            border-color: #4a4a4a !important;
        }

        .content blockquote {
            background: linear-gradient(135deg, #2d2d2d 0%, #3a3a3a 100%) !important;
            border-left-color: #f8842c !important;
        }
    }
</style>

<!-- Outlook Classic Fixes (only for MSO) -->
<!--[if mso]>
<style type="text/css">
    /* Force Light Mode for Outlook Classic */
    body { background-color: #f3f5f6; }
    .email-container { background-color: #ffffff; }
    .header { background-color: #f8f9fa; }
    .footer { background-color: #f8f9fa; }
    .content { background-color: #ffffff; }

    /* MSO Table Fixes */
    table { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }

    /* MSO Line Height Fix */
    .content p, .content li { mso-line-height-rule: exactly; }

    /* MSO List Fixes - Use standard specificity */
    .content ul {
        list-style-type: disc;
        margin-left: 0;
        padding-left: 20px;
    }
    .content ol {
        list-style-type: decimal;
        margin-left: 0;
        padding-left: 20px;
    }
    .content li {
        margin-left: 0;
        padding-left: 8px;
    }

    /* Logo Display for Classic */
    .footer .logo-dark { display: block !important; }
</style>
<![endif]-->
</head>
<body>
    <!--[if mso]>
    <v:background xmlns:v="urn:schemas-microsoft-com:vml" fill="t">
        <v:fill type="tile" color="#f5f5f5"/>
    </v:background>
    <![endif]-->
    <div class="email-container">
        <div class="header">
        </div>

        <div class="content">

            $($HtmlContent)

            <div class="tenant-info">
                <strong>Tenant:</strong> $($TenantDisplayName)<br>
                <strong>Generated:</strong> $([System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'; Get-Date -Format "dddd, MMMM d, yyyy HH:mm") <br>
                <strong>Report Version:</strong> $($ReportVersion)
            </div>

            $(if (@($Attachments).Count -gt 0) {
            @"

            <div class="attachments">
                <h3>Attached Files</h3>
                <div class="attachment-list">
                    $(($Attachments | ForEach-Object { "<div class='attachment-item'>$(Split-Path $_ -Leaf)</div>" }) -join "`n                    ")
                </div>
                <p><strong>Note:</strong> The attachments contain additional information from the generated report and can be used for more in-depth analysis.</p>
            </div>
"@
            })
        </div>

        <div class="footer">
            <div class="logo-container">
                <img class="logo-dark" alt="RealmJoin logo for dark mode" src="data:image/svg+xml;base64,$($plainBase64Logo_dark)" />
            </div>
            <div class="tagline">Companion to Intune – Application Lifecycle & Management Automation Platform</div>
            <div class="links">
                <a href="https://www.realmjoin.com">www.realmjoin.com</a> |
                <a href="https://docs.realmjoin.com">Documentation</a>
            </div>
        </div>
    </div>
</body>
</html>
"@
}

function Get-MimeTypeFromExtension {
    <#
        .SYNOPSIS
        Returns the MIME type for a given file extension.

        .DESCRIPTION
        Maps common file extensions used for tenant data exports to their appropriate MIME types.
        Supports CSV, Excel, JSON, XML, TXT, and other common formats.

        .PARAMETER FilePath
        The file path to determine the MIME type for.

        .EXAMPLE
        PS C:\> Get-MimeTypeFromExtension -FilePath "C:\temp\report.csv"
        Returns: text/csv

        .EXAMPLE
        PS C:\> Get-MimeTypeFromExtension -FilePath "C:\temp\data.xlsx"
        Returns: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet

        .OUTPUTS
        System.String. Returns the MIME type string.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()

    $mimeTypes = @{
        '.csv'  = 'text/csv'
        '.txt'  = 'text/plain'
        '.json' = 'application/json'
        '.xml'  = 'application/xml'
        '.xlsx' = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        '.xls'  = 'application/vnd.ms-excel'
        '.pdf'  = 'application/pdf'
        '.zip'  = 'application/zip'
        '.html' = 'text/html'
        '.htm'  = 'text/html'
        '.log'  = 'text/plain'
        '.md'   = 'text/markdown'
    }

    if ($mimeTypes.ContainsKey($extension)) {
        return $mimeTypes[$extension]
    }
    else {
        # Default to binary stream for unknown types
        return 'application/octet-stream'
    }
}

function Send-RjReportEmail {
    <#
        .SYNOPSIS
        Sends a RealmJoin-branded HTML email (converted from Markdown) via Microsoft Graph.

        .DESCRIPTION
        Send-RjReportEmail builds an HTML email from Markdown content, inlines a RealmJoin-styled HTML template (including light/dark logos), attaches optional files, and sends the message using the Microsoft Graph API (Invoke-MgGraphRequest).

        .PARAMETER EmailFrom
        The sender user id (user principal name or id) used for the Graph /users/{id}/sendMail call.

        .PARAMETER EmailTo
        Recipient email address(es). Can be a single address or multiple comma-separated addresses (string).
        The function sends individual emails to each recipient for privacy reasons.
        Whitespace and empty entries are automatically removed.

        .PARAMETER Subject
        Subject line for the email message.

        .PARAMETER MarkdownContent
        Report content in Markdown format. The function performs a lightweight conversion of Markdown
        to HTML and places the result into the themed HTML template used for the email body.

        .PARAMETER Attachments
        Optional array of file paths to include as attachments. Files that exist will be read,
        base64-encoded and included as file attachments. Missing files are logged and skipped.

        .PARAMETER TenantDisplayName
        Optional display name for the tenant/organization that will be shown in the email footer
        and tenant info box.

        .PARAMETER ReportVersion
        Optional string describing the report version. Will be shown in the tenant-info block.

        .EXAMPLE
        PS C:\> Send-RjReportEmail -EmailFrom "reports@contoso.com" -EmailTo "alice@contoso.com" -Subject "Weekly Report" -MarkdownContent "# Hello`nReport body..."

        .EXAMPLE
        PS C:\> Send-RjReportEmail -EmailFrom "reports@contoso.com" -EmailTo "alice@contoso.com, bob@contoso.com, team@contoso.com" -Subject "Inventory" -MarkdownContent (Get-Content .\report.md -Raw) -Attachments @('C:\temp\report.csv') -TenantDisplayName 'Contoso Ltd' -ReportVersion 'v1.2.3'

        .INPUTS
        None. All parameters are provided as arguments; this function does not accept pipeline input.

        .OUTPUTS
        None. The function sends email and writes verbose/log messages. On failure it throws an exception.

    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailFrom,

        [Parameter(Mandatory = $true)]
        [string]$EmailTo,

        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$MarkdownContent,

        [string[]]$Attachments = @(),

        [bool]$saveToSentItems = $true,

        [string]$TenantDisplayName,

        [string]$ReportVersion
    )

    # Parse and clean email addresses from EmailTo parameter
    # Split by comma, trim whitespace, remove empty entries
    # Important: Wrap in @() to ensure we always get an array (StrictMode-friendly)
    $emailRecipients = @(
        $EmailTo -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    )

    if ($emailRecipients.Count -eq 0) {
        throw "No valid email recipients found in EmailTo parameter."
    }

    Write-RjRbLog -Message "Parsed $($emailRecipients.Count) recipient(s) from EmailTo parameter" -Verbose

    # Convert Markdown to HTML using helper function
    $htmlContent = ConvertFrom-MarkdownToHtml -MarkdownText $MarkdownContent

    Write-RjRbLog -Message "Successfully converted Markdown content to HTML" -Verbose

    # Prepare email parameters
    $emailAttachments = @()
    $validatedAttachments = @()
    foreach ($file in $Attachments) {
        if (Test-Path $file) {
            $contentBytes = [IO.File]::ReadAllBytes($file)
            $content = [Convert]::ToBase64String($contentBytes)
            $mimeType = Get-MimeTypeFromExtension -FilePath $file
            $emailAttachments += @{
                "@odata.type"  = "#microsoft.graph.fileAttachment"
                "name"         = (Split-Path $file -Leaf)
                "contentType"  = $mimeType
                "contentBytes" = $content
            }
            $validatedAttachments += $file
            Write-RjRbLog -Message "Added attachment: $(Split-Path $file -Leaf) (MIME type: $mimeType)" -Verbose
        }
        else {
            Write-RjRbLog -Message "Attachment file not found: $file" -Verbose
        }
    }

    $htmlBody = Get-RjReportEmailBody -Subject $Subject -HtmlContent $htmlContent -Attachments $validatedAttachments -TenantDisplayName $TenantDisplayName -ReportVersion $ReportVersion

    # Send individual emails to each recipient for privacy
    $successfulSends = 0
    $failedSends = 0
    $failedRecipients = @()

    foreach ($recipient in $emailRecipients) {
        try {
            Write-RjRbLog -Message "Sending email to: $recipient" -Verbose

            $message = @{
                subject      = $Subject
                body         = @{
                    contentType = "HTML"
                    content     = $htmlBody
                }
                toRecipients = @(
                    @{
                        emailAddress = @{
                            address = $recipient
                        }
                    }
                )
            }

            if ($emailAttachments.Count -gt 0) {
                $message.attachments = $emailAttachments
            }

            # Send via Graph API
            #$body = @{ message = $message; saveToSentItems = $true } | ConvertTo-Json -Depth 10
            #$Uri = "https://graph.microsoft.com/v1.0/users/$($EmailFrom)/sendMail"
            #Invoke-MgGraphRequest -Uri $Uri -Method POST -Body $body -ContentType "application/json" -ErrorAction Stop
            $Resource = "/users/$($EmailFrom)/sendMail"
            $body = @{ message = $message; saveToSentItems = $saveToSentItems }
            Invoke-RjRbRestMethodGraph -Resource $Resource -Method POST -Body $body -ErrorAction Stop

            Write-RjRbLog -Message "Email sent successfully to $recipient" -Verbose
            $successfulSends++
        }
        catch {
            $failedSends++
            $failedRecipients += $recipient
            Write-RjRbLog -Message "Failed to send email to ${recipient}: $($_.Exception.Message)" -Verbose
            Write-Error "Failed to send email to ${recipient}: $($_.Exception.Message)" -ErrorAction Continue
        }
    }

    # Summary logging
    Write-RjRbLog -Message "Email sending completed: $successfulSends successful, $failedSends failed out of $($emailRecipients.Count) total recipient(s)" -Verbose

    if ($failedSends -gt 0) {
        $failedList = $failedRecipients -join ", "
        Write-RjRbLog -Message "Failed recipients: $failedList" -Verbose

        if ($successfulSends -eq 0) {
            throw "Failed to send email to all recipients: $failedList"
        }
        else {
            Write-Warning "Some emails failed to send. Failed recipients: $failedList"
        }
    }
}