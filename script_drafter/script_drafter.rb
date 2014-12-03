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
require_relative 'dialog_boxes'  #TODO: have a distribution folder set up

$debug = false

if $debug
  puts "debugging mode, press any key to continue."
  gets
end

#$PATH="C:/Users/"+username+"/Google Drive/Knowledge Engineering/Lessons - Basic IV/026/drafter/"  #path to the files

$borders="<w:tblBorders><w:top w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/><w:left w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/><w:bottom w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/><w:right w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/><w:insideH w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/><w:insideV w:val=\"single\" w:space=\"0\" w:color=\"auto\" w:sz=\"4\"/></w:tblBorders>"

$WARNING=0x00000000|0x00000010
$YESNOEX=0x00000004|0x00000030
$YESNO=0x00000004

ext=".docx"
os="_OS"
f_template="template"  #empty template
f_OS_template="OS_template" #empty OS template
f_pieces="pieces"  #document pieces
f_plan=""  #the actual lesson plan, must be in .docx format

=begin
pieces:

anote=0:note
$tutor=1:tutor line
$weak=2:weak line
$average=3:average line
$strong=4:strong line
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

$tutors=["angela","martin","stephanie"]
$stage_names=["introduction", "homework", "mental", "warm", "material", "practice", "review", "conclusion", "quiz", "exam", "stage"]

$tag={"new"=>"$new","cut"=>"$cut","weak"=>"$weak","average"=>"$average","strong"=>"$strong","tutor"=>"$tutor", "screen"=>"$screen", "insb"=>"$insb", "insa"=>"$insa",  "OS"=>"$OS", "stage"=>"$stage", "branch"=>"$branch", "endbranch"=>"$endbranch","end"=>"$end"} #insb and insa are forward- and reverse-scaffolding
$type=["$note","$line","$line","$line","$line","$answer", "$nothing", "$bb", "$line","$line", "$line","$line","$answer", "$nothing", "$stage", "$OS","$OS", "$OS", "$OS", "$note", "$CMR"]  #key strings to replace, indices following the same scheme as above!

$br="In-script branching begins."
$endbr="In-script branching ends."
$others=""  #not sure if it is necessary to leave a note about students not mentioned, probably not necessary.
$sib=", skip if behind"
$tutor=""

$item=0
$preamble=Hash.new("???")
$adv=["0","5","8"]
$grp={""=>["All"],"-"=>["All"],"w"=>["Weak"], "a"=>["Average"], "s"=>["Strong"], "wa"=>["Weak", "Average or Strong"], "ws"=>["Weak or Average","Strong"], "as"=>["Average", "Strong"], "was"=>["Weak","Average","Strong"]}


$resources=["OS_template.docx", "pieces.docx", "responses.csv", "template.docx"]
$RUNPATH=Dir.pwd+"/"
$PATH=$RUNPATH
username=ENV["USERNAME"]

if username=="yliu" && $debug
  $RUNPATH="C:/Users/yliu/SkyDrive/RM-synced/cogitatio/script_drafter/"
  $PATH=$RUNPATH
  Dir.chdir($RUNPATH)
end


=begin
=end

if !(Dir.glob("*").include?("resources"))
  show("Folder \\resources\\ missing, exiting.","WARNING",$WARNING)
  exit
end

arg=ARGV[0].to_s  #get path from command line input

begin  
  if Dir.exists?(arg)
    $PATH=arg
    arg=""
  else
    puts arg+" does not exist." if arg!=""
    $PATH=getfolder("Select lesson folder containing lesson plan.")
  end
  
  if $PATH=="" || $PATH.nil?
    puts "Exiting."
    exit
  end
  
  $PATH=$PATH.gsub("\\","/")
  $PATH=$PATH+"/" unless $PATH.match(/[\/]$/)
  
  $PATH=$RUNPATH if username=="yliu" && $debug
  Dir.chdir($PATH) 
  list=Dir.glob("*[0-9][0-9][0-9]*.docx") & Dir.glob("*plan*.docx")
  if list==[]
    show("Lesson plan must be in docx format, has a 3 digit lesson number and \"plan\" in its name.", "Lesson plan not found.", $WARNING)
  else
    puts "Multiple lesson plan files found." if list.size>=2
    begin
      list.each do |f|
        rc=showmessage("Is this the correct lesson plan file? \n\n     #{f}","Confirm",$YESNO)
        if rc==6
          f_plan=f.split(".")[0]
          break
        elsif rc==7 && list.size==1
          puts "Exiting."
          exit
        end
      end
    end until f_plan!=""
  end
end until f_plan!=""

$lesson=/(\d{3})/.match(f_plan)[1]
puts "Lesson number: "+$lesson

=begin
=end

print "Getting resources... "

Dir.chdir($RUNPATH+"resources/")
list=Dir.glob("*.*")

$resources.each do |f|
  if !(list.include?(f))
    show("#{f} missing, exiting.","WARNING",$WARNING)
    exit
  end
end

$pcs=Array.new
$cmt=nil

fs=DOCX.open(f_pieces+ext)

fs.getpcs 

$resp=Hash.new
$resp_count=Hash.new
f=File.open("responses.csv","r")
f.readline
f.each do |line|
  a=line.strip.split(",")
  name=a[0].strip.downcase
  if $resp[name].nil?
    $resp[name]=Array.new
    $resp_count[name]=Hash.new
  end
  
  a[2].to_i.times.each do 
    $resp[name] << a[1].strip 
  end
  $resp_count[name][a[1].strip] = [a[3].to_i,0]  #cap, count
  
end

fos=DOCX.open($RUNPATH+"resources/"+f_OS_template+ext)

puts "Done."

=begin
=end


$content=Hash.new
$os_content=Array.new
$comments=Hash.new
$os_stuff=Array.new
$plan_items=Array.new
$os_items=Array.new
$submit_count=Hash.new(0)
$word_count=Hash.new(0)
$chr=nil
$draft_path=$PATH+$lesson+"_script_drafts/"

Dir.chdir($PATH) 

fos_name=$lesson+os+ext
$os_exists=(File.exists?(fos_name))? 1:0

#fp=DOCX.open(f_plan+ext)

if $os_exists==0 #|| $debug
  rc=show("OS file #{fos_name} not found, generate OS draft?","Generate OS?",$YESNO)
  if rc==7
    puts "Exiting"
    exit
  elsif rc==6
    print "Generating #{fos_name}... "
    fp=DOCX.open(f_plan+ext)
    fp.extract_cells1
    fos.add_os
    fos.save($PATH+fos_name)
    $os_exists=1
    puts "Done."
    
    $os_content.each do |a|
      if a[0]!=0
        $os_items<<a[0]
      end
    end
    
    print "OS file #{fos_name} generated with lesson items:\n(Branchings appear as duplicates.)\n\n"
    print $os_items
    puts
    
    rc=show("Continue?","Continue?",$YESNO)
    
    if rc==7
      puts "Exiting"
      exit
    end
    
  end
end

if $os_exists==1
  rc=show("OS file #{fos_name} found, generate script drafts?","Generate scripts?",$YESNO)
  if rc==7
    puts "Exiting"
    exit
  elsif rc==6
  
    puts "Generating script drafts... "
    fp=DOCX.open(f_plan+ext)
    fp.extract_cells1
    fos=DOCX.open($PATH+fos_name)
    fos.extract_os
    
    ($os_items-$plan_items).each do |num| 
      $content[num]=Array.new
    end

    if Dir.exists?($draft_path)
      rc=showmessage("Script drafts folder already exists: \n\n#{$draft_path} \n\nProceed?  (Files WILL be overwritten.)","WARNING",$YESNOEX)
      
      if rc==7
        puts "Exiting."
        exit
      end
    else
      puts "Creating script drafts folder: #{$draft_path}"
      Dir.mkdir $draft_path
    end
    
    ft=Array.new
    $os_items.uniq!
    $os_items.each do |num|
      
      nm="000"[0..-(num.to_s.length+1)]+num.to_s

      ft[num]=DOCX.open($RUNPATH+"resources/"+f_template+ext)
      ft[num].add(num)
      ft[num].save($draft_path+$lesson+"-"+nm+ext)
      puts "script file "+$lesson+"-"+nm+ext+" generated."

    end
    
    
    
  end
end

puts "Exiting."
exit


####





