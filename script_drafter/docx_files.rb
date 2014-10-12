
class DOCX
  def self.open(path)
    self.new(path)
  end


  def initialize(path)
    @replace = {}
    @zip = Zip::ZipFile.open(path)
    
    @doc=Nokogiri::XML(@zip.read("word/document.xml")) {|x| x.noent}

    if @zip.find_entry("word/comments.xml")
      @cmt=Nokogiri::XML(@zip.read("word/comments.xml")) {|x| x.noent}
      
    else
      @cmt=nil
    end
  end

  attr_accessor :doc, :cmt

  def save(path)
    Zip::ZipFile.open(path, Zip::ZipFile::CREATE) do |output|
      @zip.each do |entry|
        output.get_output_stream(entry.name) do |o|
          if @replace[entry.name]
            o.write(@replace[entry.name])
          else
            o.write(@zip.read(entry.name))
          end
        end
      end
    end
    @zip.close
  end
  
  def add(numm)


    cmt_count=0 #cdoc.at(".//comment")["w:id"].to_i

    
    itm="000"[0..-(numm.to_s.length+1)]+numm.to_s
    
    @doc.xpath("//w:p").each do |t|
    
      if t.inner_html.include?("$script")

        if $os_stuff[numm][0][1]!=""

          str=itm+" ("+$os_stuff[numm][0][1].content.strip+")"
        else
          str=itm
        end

        #print str
        #gets
        
        t.rep("$script",str)

      end
    
      #print $preamble
      #gets
      
      $preamble.each do |k,v|
        t.rep("$"+k,v)
      end

    end
    
    @doc.at(".//w:commentRangeStart")["w:id"]=cmt_count.to_s
    @doc.at(".//w:commentRangeEnd")["w:id"]=cmt_count.to_s
    @doc.at(".//w:commentReference")["w:id"]=cmt_count.to_s
    
    @cmt.rep("$comment",$os_stuff[numm][-1][1].content.strip)
    

    @cmt.at(".//w:comment")["w:id"]=cmt_count.to_s
    @cmt.at(".//w:comment")["w:author"]="KE"
    @cmt.at(".//w:comment")["w:date"]=DateTime.now.to_s
    @cmt.at(".//w:comment")["w:initials"]="KE"

    
    
    @doc.xpath(".//w:p").each do |p|
      if p.content.strip.match(/\$char(\d{2})/)
        p.content.strip.scan(/\$char(\d{2})/).each do |x|
          
          #print $chr.content.strip
          #gets
          
          i=x[0][0].to_i
          j=x[0][1].to_i
          str=$chr.rows[i].cells[j].content.strip
          p.rep("$char"+x[0],str)
        end
      end
    end
    
    ary=$os_stuff[numm][1..-2]+$content[numm] #adding os stuff, then copy from lesson plan
    

    prev_rsp=Array.new  #array to record previous responses to avoid
    prev_rsp[5]=""
    prev_rsp[12]=""
    
    ary.each do |aa|

      i=aa[0] 
      cnt=aa[1] #content node 
      
      if cnt.class!=String && cnt!=nil
        cnt.strike!.remove_comments
      end
      
      rr=/\[.{5,}?\]/
      
      cnt1=cnt.content.gsub(/\$new\(.{,3}\)/,"")

      
      $tag.values.each do |tg|
        cnt1=cnt1.gsub(tg,"").strip.gsub(/\s/," ").squeeze(" ")
        cnt.rep(tg,"") if cnt!=nil

      end   
      
      cnt2=cnt1
      
      $pcs[i].each do |m| #each pcs[i] is an array of nodes
   
        #HERE!!!
        nn=m.dup

        if nn.inner_html.include?($type[i]) || nn.inner_html.include?("$right.") #ugly temp fix
          
          if i==7 && cnt.xpath(".//w:drawing").to_a==[] #screen, no drawing

            rto=nn.at(".//w:r")
            
            cnt.children.to_a.reverse.each do |p|
              if p.name=="p"  || p.name=="tbl"
              
                p.remove_comments
                
                p.add_onscr
                
                p.add_borders if p.name=="tbl"
  
                rto.after(p)
              end
            end
            
            nn.rep($type[i],"").rep(rr,"")
            
          elsif i==7 && cnt.xpath(".//w:drawing").to_a!=[]  #drawing
            tg=" [DRAWING]" 

            nn.rep($type[i],cnt1.gsub(rr,"")+tg)
          
          
          else
            
            mtch=cnt1.scan(rr)
            mtch.each do |m|
              cnt2=cnt1.gsub(m,"")
            end 

            nn.rep($type[i],cnt2.gsub(rr,""))
            
            if i== 5 || i== 12 #add tutor's response
              
              if !($resp[$tutor].nil?)

                begin
                  rsp = $resp[$tutor].sample
                end until ($resp_count[$tutor][rsp][1]<$resp_count[$tutor][rsp][0] && !(prev_rsp.include?(rsp)))
                $resp_count[$tutor][rsp][1]+=1  #count the times rsp is used
                prev_rsp[i]=rsp

              else
                rsp="That's right."
              end
              
              nn.rep("$right.",rsp)
            end
            
          end 

          if cnt1.match(rr)  #add comments
          
            cnt1.scan(rr).each do |tcmt|
              cmt_count+=1
              
              nn.search(".//w:r").to_a[0].comment_start(cmt_count.to_s)
              nn.search(".//w:r").to_a[-1].comment_end(cmt_count.to_s)           
              
              @cmt.add_comment(cmt_count.to_s, tcmt)
            end
          end

                   
 
        end  
         
        @doc.root.child<<(nn)
      end
    
    end
    @replace["word/document.xml"] = @doc.to_xml
    @replace["word/comments.xml"] = @cmt.to_xml
  

  end
 
  def add_os

    $preamble.each do |k,v|
      @doc.rep("$"+k,v)
    end

    $os_content_merged=Array.new

    
    temp=nil
    
    $os_content.each do |aa|
    
      if aa[1].class!=String && aa[1]!=nil
        aa[1].strike!

      end
    
      if aa[0]==temp
      
        if $os_content_merged[-1][1]==""
          $os_content_merged[-1][1]+=aa[1].content
        else
          $os_content_merged[-1][1]+="<br><br>"+aa[1].content
        end
        
        $os_content_merged[-1][2]+=aa[2]
      else
        $os_content_merged << [aa[0], aa[1].content, aa[2]]
      end
      
      temp=aa[0]
    
    end
    
    
    $os_content_merged.each do |aa|
    
    
      itm=aa[0].to_s
      cnt1=aa[1] # string
      par=aa[2]  #parameter

      #ncol=(par.length==0)? 1 : par.length
      ncol=par.length
      if itm=="0"
        ncol=-1
      end
      
      cnt1=cnt1.gsub(/\$new\(.*\)/,"")
      $tag.values.each do |tg|
        cnt1=cnt1.content.gsub(tg,"")
      end
      
      groups=Array.new
      
      $grp[par.downcase].each do |g|
        groups<<g
      end
      
      it=Array.new
      (0..2).each do |i|
        it[i]=itm[0..(-2)]+$adv[i]
      end
      
      if ncol==0
        
      else
        (0...ncol).each do |i|
          if par[i].upcase==par[i] && par!=""
            groups[i]+=$sib
          end
          
        end
      end
      
      
      $pcs[15+ncol].each do |m| #pcs[i] is an array of nodes

        nn=m.dup

        
        (nn/".//w:t").each do |t| #figure out how missing this "." messed the whole thing up

          t.rep($tag["OS"],cnt1)
          cc=t.inner_html.match(/\$Group(\d{1})/)[1].to_i
          t.rep("$Group"+cc.to_s,groups[cc])
          
          cc=t.inner_html.match(/\$item(\d{1})/)[1].to_i
          t.rep("$item"+cc.to_s,it[cc])

          t.rep("$type"+cc.to_s, "Together") 

          t.rep("$COMMENT", $submit_count[itm.to_i].to_s+" submits "+$word_count[itm.to_i].to_s+" words") 

          t.rep("$stage", cnt1) 

        end

        @doc.root.child<<(nn) 
      end
    
    end
    @replace["word/document.xml"] = @doc.to_xml
    
  

  end
 
 
  def getpcs
  
    i=-1
    
    @doc.at(".//w:body").children.each do |node|
      if node.content.strip.include?("##")  #catch strings after ##

        i+=1
        $pcs[i]=Array.new
      else
        $pcs[i] << node

      end
      
    
    end
    
    $cmt=@cmt.at(".//w:comment").dup

  end
  

 def expose
  
    xml = @zip.read("word/comments.xml")
    doc = Nokogiri::XML(xml) {|x| x.noent}

    puts doc.name
    puts doc.child.name
    puts doc.child.child.name
    puts doc.child.child.children.to_a.size
    puts doc.child.child.children.to_a.map{|x| x.name+"->"+x.content}
    puts
    
    
    gets
    
    doc.xpath(".//w:tbl").each do |n|
      puts n.name
      gets
    end
    
    gets
    
    (doc/"//w:tbl").each do |t|
    #doc.child.child.children.each do |n|
      
      (t/".//w:tc").each do |n|
      puts n.name
      puts n.content.strip
      #puts n.inner_html
      gets
      end
    end
  end


  def extract_os

    pattern=[0,0,19,0,19] #[0,0,7,0,19]  #for the 5 rows in each table
    
    @doc.at(".//w:tbl").remove #remove the "file team" table
    
    $chr=@doc.at(".//w:tbl").dup #copying the characters table
    
    $tutor=$chr.rows[0].cells[1].content.strip.downcase #getting tutor's name
    
    if !($tutors.include?($tutor))
      rc=showmessage("Incorrect tutor name: \"#{$chr.rows[0].cells[1].content.strip}\".  Proceed? \n (All correct responses will be set to \"That's right.\")","WARNING",$WARNING)
      if rc==7
        puts "Exiting."
        exit
      end
    end
    
    @doc.at(".//w:tbl").remove  #remove the characters table
    
    @doc.at(".//w:body").children.each do |tbl|  
    
      next if tbl.name!="tbl"
      
      oscells=Hash.new
      branches=Array.new 
      rows=tbl.rows

      nrow=rows.size
      
      (rows[1].xpath(".//w:tc").to_a+rows[0].xpath(".//w:tc").to_a).each do |cell|
      
        if cell.content.strip.match(/^\d{3}/)
          lin=cell.content.strip.match(/^(\d{3})/)[1].to_i

          $os_stuff[lin]=Array.new
          $os_items << lin 
          branches << lin

        end
      end
      
      #useful debugging: when the code fails to catch branches ("no implicit conversion from nil to integer in []")
      #print branches
      #gets
      
      next if branches==[]  #danger!

      i=5-nrow  #bad, expected row index to copy os stuff from
      if i==1
        branches.each do |t|
          $os_stuff[t]<<[0,""]  #this takes care of the case when the table does not have the strength row
        end
      end

      rows.each do |row|
        j=0 #column index in the table
        
        #print branches
        #gets
        
        unless i==1
          row.cells.each do |cell|
            
            if cell.content.strip!="$DESIGN" && cell.content.strip!=""
              $os_stuff[branches[j]] << [pattern[i],cell]
  
            end

            j+=1
          
          end
        end
        i+=1

      end
    end
  end
  
  def extract_cells1
    
    if !(@cmt.nil?)
      @cmt.xpath(".//w:comment").each do |cr|
      
        $comments[cr["w:id"].to_s.strip]=cr.content.strip

      end
    end
    
    temp=Array.new
    row=Array.new
    
    parb=""
    
    tb0=@doc.at(".//w:tbl")
    ary=tb0.rows.map{|x| x.cells.map{|y| y.content.strip}}
    $preamble=Hash[ary.map{|p| [p[0].downcase,p[1]]}]
    
    #print $preamble
    #gets
    
    $preamble.delete("lesson")
    $preamble["lesson_num"]=$lesson
    $preamble["lesson_type"]=$preamble["lesson type"]
    $preamble["teacher"]=$preamble["author"]
    
    #print $preamble
    #gets

    tb0.remove
    
    @doc.at(".//w:body").children.each do |n|
      
      if n.name=="p" && n.content.downcase.stage?

        $item=100*(1+$item/100)
        $item+=10 #a new stage carries implicit $new
        $content[$item]=Array.new
        $plan_items << $item
        
        $os_content << [0,n,""]
        $os_content << [$item,"",""] #dummy OS item

      elsif  n.name=="tbl"
        
        ff=0
        flag=0
        fftemp=Array.new
        target=Array.new
        
        therows=n.rows
        
        therows[1..-1].each do |r|
          
 
          row=r.cells
          
          
          if r.content.include?($tag["insa"])
            flag=7
          end
          
          if r.content.include?($tag["insb"])
            flag=7
            ff=1

          end
          
          if row[0].content.include?($tag["new"]) 
            
            if row[0].content.match(/\$new\([^()]{,3}\)/)
              par=row[0].content.match(/\$new\(([^()]{,3})\)/)[1]
            else
              par=""
            end
            
            if therows.index(r)>=2  #$new is ignored if in the first content row (rn=2) of stage
              $item+=10
              $content[$item]=Array.new
              $plan_items << $item
            end
            
            #$os_content<<[$item,row[0],""] #$new carries implicit $OS
            
            $os_content.add_new([$item,"",par])
            
          end
          
          if row[0].content.include?($tag["OS"])
            $os_content.add_new([$item,row[0],""])
          end
          
          if row[1].content.include?($tag["OS"])
            $os_content.add_new([$item,row[1],""])
          end
            
            
        #adding to the content array:
          
          
          if ff==1
            target=fftemp
          else
            target=$content[$item]
          end
          
          
          if r.content.include?($tag["branch"])           
            if r.content.match(/\$branch\(.{,3}\)/)
            

              parb=r.content.match(/\$branch\((.{,3})\)/)[1]
              
              target << [0,$br]
              
              
              str=$grp[parb.downcase][0]

              if parb[0]==parb[0].upcase
                str+=$sib
              end
              
              target << [0,str+":"]
              
              

              
            end
          end
          

          
          
          unless row[1].content.include?($tag["cut"]) || row[1].strike!.content.strip==""
          

            row[1].xpath(".//w:commentReference").each do |cr|
              if target.nil? 
                print row[1].content
                gets
              end
              target << [19, $comments[cr.attribute("id").to_s.strip]]
              
            end
            
            target << [1+flag,row[1]]
            
 
            $word_count[$item]+=row[1].content.split(" ").size
          end
          
          if row[0].content.include?($tag["screen"])
          
            row[0].xpath(".//w:commentReference").each do |cr|
              target << [19, $comments[cr.attribute("id").to_s.strip]]
            end
          
            target << [7,row[0]]
          end
          
          unless row[2].content.include?($tag["cut"]) || row[2].strike!.content.strip=="" || row[1].content.include?($tag["cut"]) || row[1].strike!.content.strip==""
              row[2].xpath(".//w:commentReference").each do |cr|

              target << [19, $comments[cr.attribute("id").to_s.strip]]
            end
          
            if row[2].content.include?($tag["tutor"])
              target << [1+flag,row[2]]
              $word_count[$item]+=row[2].content.split(" ").size
            elsif row[2].content.include?($tag["weak"]) 
              target << [2+flag,row[2]]
              $word_count[$item]+=row[2].content.split(" ").size
            elsif row[2].content.include?($tag["average"]) 
              target << [3+flag,row[2]]
              $word_count[$item]+=row[2].content.split(" ").size
            elsif row[2].content.include?($tag["strong"]) 
              target << [4+flag,row[2]]
              $word_count[$item]+=row[2].content.split(" ").size
            else
              target << [5+flag,row[2]]
              $word_count[$item]+=row[2].content.split(" ").size
              #if (flag==0 && ff!=2) || ff==1 || (flag==7 && ff!=1)
              #  target << [6+flag,""]
              #end
              $submit_count[$item]+=1
            end
          end
          
          

          if r.content.include?($tag["endbranch"]) && parb!=""        

            (1...($grp[parb.downcase].size)).each do |i|
            
              str=$grp[parb.downcase][i]

              if parb[i]==parb[i].upcase
                str+=$sib
              end
              
              target << [0,str+":"]
            end
            

            target << [0,$endbr]
            parb=""
            

          end

          #above: adding content
          
          if ff==2
            ff=0
            $content[$item]+=fftemp
            fftemp=Array.new 
          end
          
          if r.content.include?($tag["end"])

            flag=0
            if ff==1
              ff=2
            end

          end
          

        end
      end
      
    
    end
    
 
  end
  
  
 
end
