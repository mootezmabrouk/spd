Add-Type -AssemblyName System.IO.Compression.FileSystem

function Update-PptxText($zipPath, $outPath) {
    Copy-Item $zipPath $outPath -Force
    $zip = [System.IO.Compression.ZipFile]::Open($outPath, [System.IO.Compression.ZipArchiveMode]::Update)
    $entry = $zip.GetEntry("ppt/slides/slide1.xml")
    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream)
    $xmlContent = $reader.ReadToEnd()
    $stream.Close()
    $entry.Delete()

    # Replacements
    $xmlContent = $xmlContent -replace "Montassar", "Mootez MABROUK"
    $xmlContent = $xmlContent -replace "1\. Contexte &amp; Objectifs", "Introduction"
    $xmlContent = $xmlContent -replace "1\. Contexte & Objectifs", "Introduction"
    $xmlContent = $xmlContent -replace "2\. Données &amp; Méthodologie", "Méthodologies"
    $xmlContent = $xmlContent -replace "2\. Données & Méthodologie", "Méthodologies"
    $xmlContent = $xmlContent -replace "3\. Résultats Expérimentaux", "Résultats"
    $xmlContent = $xmlContent -replace "4\. Conclusion &amp; Perspectives", "Conclusion"
    $xmlContent = $xmlContent -replace "4\. Conclusion & Perspectives", "Conclusion"
    $xmlContent = $xmlContent -replace "Références &amp; Remerciements", "Références"
    $xmlContent = $xmlContent -replace "Références & Remerciements", "Références"
    $xmlContent = $xmlContent -replace "Projet de Recherche &amp; PFE · École Nationale des Sciences de l'Informatique \(ENSI\), Univ\. de la Manouba", "Laboratoire CRISTAL, Groupe de recherche GRIFT, École Nationale des Sciences de l'Informatique, Université de Manouba"
    $xmlContent = $xmlContent -replace "Projet de Recherche & PFE · École Nationale des Sciences de l'Informatique \(ENSI\), Univ\. de la Manouba", "Laboratoire CRISTAL, Groupe de recherche GRIFT, École Nationale des Sciences de l'Informatique, Université de Manouba"

    $newEntry = $zip.CreateEntry("ppt/slides/slide1.xml", [System.IO.Compression.CompressionLevel]::Optimal)
    $writer = New-Object System.IO.StreamWriter($newEntry.Open())
    $writer.Write($xmlContent)
    $writer.Close()
    $zip.Dispose()
    Write-Host "Updated $outPath successfully."
}

Update-PptxText "reports/poster_spd_ensi_modele_60x100cm.pptx" "reports/poster_spd_ensi_modele_60x100cm.pptx"
Update-PptxText "reports/poster_spd_ensi_modele_60x100cm.pptx" "Modele_Poster_ENSI_60x100cm (1).pptx.pptx"
Update-PptxText "reports/poster_spd_ensi_modele_60x100cm.pptx" "reports/poster_spd_ensi_60x100cm_Mootez_Mabrouk.pptx"
