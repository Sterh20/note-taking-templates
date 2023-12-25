# Define a parameter for the script that specifies the number of copies to make
param (
    [Parameter(Mandatory = $true, Position = 0)]
    [int]$copyCount
)

# Check for valid $copyCount value
if ($copyCount -le 0) {
    Write-Host "Invalid copy count. Please provide a positive integer."
    exit
}

# Get all PNG and JPG files in the current directory and sort them by the last two caracters of a file base name
$files = Get-ChildItem -Path ($PSScriptRoot + '\*') -Include *.png, *.jpg | Sort-Object { $_.BaseName[-2..-1] }

$counter = 1
$templates = @()
# Cover is an array becouse it can be not only an array, but also a number of templates, 
# that represent an agreagate or a birds aye view on other templates in the notebook (like slidesheet and slide templates)
$covers = @()

# Loop through each file and split them to cover and template
foreach ($file in $files) {
    # Find last two characters in the file's base name
    $lastTwoCharacters = $file.BaseName[-2..-1] -join ''
    # Check if a file is a cover or a template
    if ($lastTwoCharacters -match '^\d{2}$' -and $lastTwoCharacters -eq '00') {
        $covers += $file
    }
    else {
        $templates += $file
    }
}

# Rename the cover file if exists
if ($covers) {
    foreach ($cover in $covers) {
        Rename-Item -Path $cover.FullName -NewName ($cover.DirectoryName + '\' + '{0:D3}-{1}' -f $counter, $cover.Name)
        $counter++
    }
}

# Copy template files
for ($i = 0; $i -lt $copyCount; $i++) {
    foreach ($template in $templates) {
        Copy-Item -Path $template.FullName -Destination ($template.DirectoryName + '\' + '{0:D3}-{1}' -f $counter, $template.Name)
        $counter++
    }
}


# Move old templates to recycle bin
$sh = new-object -comobject "Shell.Application"

foreach($template in $templates) {
    $ns = $sh.Namespace(0).ParseName($template)
    $ns.InvokeVerb("delete")
}

Write-Host "Script completed successfully."
