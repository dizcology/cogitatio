require 'rubygems'
require 'zip/zip' # gem install zip-zip (maybe after gem install rubyzip)
require 'Date'

#$PATH="C:/Users/yliu/SkyDrive/RM-synced/cogitatio/script_drafter/"
$PATH = "C:/Users/eabalo/Desktop/script_drafter_TE/"
Dir.chdir($PATH)

system "ocra --output script_drafter_TE_beta.exe script_drafter_TE.rb"

#files=["script_drafter_beta.exe", "responses.csv", "OS_template.docx", "pieces.docx", "tags.docx", "template.docx"]

files=["script_drafter_TE_beta.exe", "README.docx", "resources/OS_template_TE.docx","resources/pieces.docx","resources/responses.csv","resources/template.docx"]

zipped="script_drafter_TE_"+DateTime.now.to_s[0..9]+".zip"

if File.exists?(zipped)
  File.delete(zipped)
end

Zip::File.open(zipped, Zip::File::CREATE) do |z|
  files.each do |f|
    z.add(f,f)
  
  end
end
