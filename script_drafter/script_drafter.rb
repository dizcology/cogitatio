require 'rubygems'
require 'zip/zip' # gem install zip-zip (maybe after gem install rubyzip)
#require 'rubyzip'
require 'nokogiri'
require 'fileutils'
require 'io/console'
require 'Date'
require_relative 'misc.rb'
require_relative 'xml_nodes.rb'
require_relative 'docx_files.rb'

$PATH=""

username=ENV['USERNAME']

$PATH="C:/Users/"+username+"/Google Drive/Knowledge Engineering/Lessons - Basic IV/031/"  #path to the files


ext=".docx"
os="_OS"

f_template="template"  #empty template
f_OS_template="OS_template" #empty OS template
f_pieces="pieces"  #document pieces
f_plan=""  #the actual lesson plan, must be in .docx format

if $PATH!=""
  Dir.chdir($PATH)
else
  $PATH=Dir.pwd+"/"
end

print "Lesson plan file name must have 3 digit lesson number.\n\n"


Dir.glob("*[0-9][0-9][0-9]*.docx").each do |fn|

  next if !(fn.match(/[P|p]lan/))
  
  print "Generate OS file for \""+fn+"\"? (y/n) "
  r=STDIN.getch
  print r
  puts
  if r.downcase=="n"
    next
  elsif r.downcase=="y"
    f_plan=fn.split(".")[0]
    break
  else
    redo
  end
  
end

if f_plan==""
  exit
end


$lesson=/(\d{3})/.match(f_plan)[1]


$pcs=Array.new
$cmt=nil

$content=Hash.new
$os_content=Array.new
$comments=Hash.new

$os_exists=(File.exists?($PATH+$lesson+os+ext))? 1:0
$os_stuff=Array.new

$plan_items=Array.new
$os_items=Array.new

$submit_count=Hash.new(0)

$word_count=Hash.new(0)

#Dir.chdir($PATH)

#ft=DOCX.open($PATH+f_template+ext)
fs=DOCX.open($PATH+f_pieces+ext)
fp=DOCX.open($PATH+f_plan+ext)


fs.getpcs 


$borders="<w:tblBorders><w:top w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/><w:left w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/><w:bottom w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/><w:right w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/><w:insideH w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/><w:insideV w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/></w:tblBorders>"



=begin
pieces:

anote=0:note
$tutor=1:tutor line
$weak=2:weak line
$average=3:average line
$strnng=4:strong line
submit=5:submit
nextb=6:next in branch
$screen=7:screen
$tutorb=8:tutor branchline
$weakb=9:weak branchline
$averageb=10:average branchline
$strongb=11:strong branchline
$submitb=12:secondary submit
$nextbb=13:next in secondary branch
stage=14:stage
$new=15:single table no borders
$new(.)=16:single branch table
$new(..)=17:double branch table
$new(...)=18:triple branch table
$ke=19:KE note style
$cmt=20:comment
=end



$tag={"new"=>"$new","cut"=>"$cut","weak"=>"$weak","average"=>"$average","strong"=>"$strong","tutor"=>"$tutor", "screen"=>"$screen", "insb"=>"$insb", "insa"=>"$insa",  "OS"=>"$OS", "stage"=>"$stage", "branch"=>"$branch", "endbranch"=>"$endbranch","end"=>"$end"} #insb and insa are forward- and reverse-scaffolding
$type=["$note","$line","$line","$line","$line","$answer", "$nothing", "$bb", "$line","$line", "$line","$line","$answer", "$nothing", "$stage", "$OS","$OS", "$OS", "$OS", "$note", "$CMR"]  #key strings to replace, indices following the same scheme as above!

$br="In-script branching begins."
$endbr="In-script branching ends."
$others=""  #not sure if it is necessary to leave a note about students not mentioned, probably not necessary.
$sib=", skip if behind"

$item=0
$preamble=Hash.new("???")
$adv=["0","5","8"]
$grp={""=>["All"],"-"=>["All"],"w"=>["Weak"], "a"=>["Average"], "s"=>["Strong"], "wa"=>["Weak", "Average or Strong"], "ws"=>["Weak or Average","Strong"], "as"=>["Average", "Strong"], "was"=>["Weak","Average","Strong"]}

#$content[$item]=Array.new

####


if $os_exists==0
  fp.extract_cells1
  fos=DOCX.open($PATH+f_OS_template+ext)
  fos.add_os
  fos.save($PATH+$lesson+os+ext)

  $os_content.each do |a|
    if a[0]!=0
      $os_items<<a[0]
    end
  end
  
  print "OS file "+$lesson+os+ext+" generated with lesson items:\n"
  print $os_items
  puts
  exit  #only generate OS

else
  
  print "#{$lesson+os+ext} exists.\n"
  print "Generate scripts file for lesson #{$lesson}? (y/n) "
  r=STDIN.getch
  print r
  puts
  
  if r.downcase=="n"
    exit
  end

  fp.extract_cells1
  
  fos=DOCX.open($PATH+$lesson+os+ext)
  fos.extract_os

end

($os_items-$plan_items).each do |num| 

  #num1=10*(num/10)
  #$content[num]=$content[num1]   #duplicate the lowest branch
  
  
  $content[num]=Array.new
  
end

$draft_path=$PATH+$lesson+"_script_drafts/"

unless File.exists?($draft_path[0..-2])

  Dir.mkdir $draft_path[0..-2]
end

ft=Array.new
$os_items.each do |num|
  
  
  
  nm="000"[0..-(num.to_s.length+1)]+num.to_s

  ft[num]=DOCX.open($PATH+f_template+ext)
  ft[num].add(num)
  ft[num].save($draft_path+$lesson+"-"+nm+ext)
  print "script file "+$lesson+"-"+nm+ext+" generated.\n"

end



