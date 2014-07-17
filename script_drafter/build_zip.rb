require 'rubygems'
require 'zip/zip' # gem install zip-zip (maybe after gem install rubyzip)
require 'Date'

$PATH="C:/Users/yliu/SkyDrive/RM-synced/cogitatio/script_drafter/"
Dir.chdir($PATH)

system "ocra --output script_drafter_beta.exe script_drafter.rb"

#files=["script_drafter_beta.exe", "responses.csv", "OS_template.docx", "pieces.docx", "tags.docx", "template.docx"]

files=["script_drafter_beta.exe", "tags.docx", "resources/OS_template.docx","resources/pieces.docx","resources/responses.csv","resources/template.docx"]

zipped="script_drafter_"+DateTime.now.to_s[0..9]+".zip"

Zip::File.open(zipped, Zip::File::CREATE) do |z|
  files.each do |f|
    z.add(f,f)
  
  end
end
