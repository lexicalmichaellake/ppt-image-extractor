#!/bin/bash
#Accepts the filename of the PowerPoint from which you want to extract the images
path="$1"
#Displays the filename chosen for confirmation
echo "$path"
#loads LibreOffice in headless mode to convert your PowerPoint into something Linux can work with
libreoffice --headless --convert-to odp "$path"
#Slices off the extension
path="${path%.*}"
#Unzips the resulting LibreOffice presentation file
unzip "${path##*.}.odp"
#puts all XML tags in the content document of the PowerPoint onto their own line to make them more searchable for our purposes
sed 's|>|>\n|g' content.xml > content_2.xml
#looks for lines with common image types and outputs those to a new file
grep -E -i "*.png|*.jpg|*.jpeg|*.bmp|*.gif|*.tif|*.svg" content_2.xml > content_3.xml
#removes the left-hand XML co-text of an image path stored in the content.xml document
sed 's|<draw.*Pictures\/|Pictures\/|g' content_3.xml > content_4.xml
#removes the right-hand XML co-text of an image path stored in the content.xml document
sed 's|\" xlink.*>||g' content_4.xml > content_5.xml
#ensures that only the lines containing the path to your pictures are kept in the next output file
grep -E -i "Pictures\/" content_5.xml > content_6.xml

# loads the file containing the paths to your image files
IFS=$'\n' read -d '' -r -a lines < content_6.xml
#Makes a folder named for the initial PPT you loaded to reduce the chances of files being overwritten if you want multiple PPTs in the same folder converted.
mkdir "Pictures_$path"
#read the generated array of filenames, parse out the extension, and append it to an auto-incrementing variable which is the new filename
for i in "${lines[@]}"; do extension="${i##*.}"; filename="${i%.*}"; newfile=$((newfile + 1)); mv $i "Pictures_$path/$newfile.$extension"; done

#cleanup phase
rm 'content.xml'
rm 'content_2.xml'
rm 'content_3.xml'
rm 'content_4.xml'
rm 'content_5.xml'
rm 'content_6.xml'
