require 'nokogiri'
require 'zip'
require 'cgi'

BODY_START = '<w:body>'
#BODY_END = '<w:sectPr w:rsidR="00D923D3">'
BODY_END = '<w:sectPr w:rsidR="00D923D3" w:rsidRPr="00F53192">'
TABLEROW_PORTS = '<w:tr w:rsidR="00F53192" w:rsidTr="00C044F6"><w:trPr><w:jc w:val="center"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="1417" w:type="dxa"/></w:tcPr><w:p w:rsidR="00F53192" w:rsidRPr="006A6D2E" w:rsidRDefault="00F53192" w:rsidP="001C2F34"><w:pPr><w:pStyle w:val="figure"/></w:pPr><w:r><w:t>MRDPORT</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1276" w:type="dxa"/></w:tcPr><w:p w:rsidR="00F53192" w:rsidRPr="006A6D2E" w:rsidRDefault="00F53192" w:rsidP="001C2F34"><w:pPr><w:pStyle w:val="figure"/></w:pPr><w:r w:rsidRPr="006A6D2E"><w:rPr><w:rtl/></w:rPr><w:t>باز</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1276" w:type="dxa"/></w:tcPr><w:p w:rsidR="00F53192" w:rsidRPr="006A6D2E" w:rsidRDefault="00F53192" w:rsidP="001C2F34"><w:pPr><w:pStyle w:val="figure"/></w:pPr><w:r><w:t>MRDNAME</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1276" w:type="dxa"/></w:tcPr><w:p w:rsidR="00F53192" w:rsidRPr="006A6D2E" w:rsidRDefault="00F53192" w:rsidP="001C2F34"><w:pPr><w:pStyle w:val="figure"/><w:rPr><w:rtl/></w:rPr></w:pPr><w:r><w:t>MRDPROD</w:t></w:r></w:p></w:tc></w:tr>'

def x bind
		require 'readline'
        puts "Execuation mode: "
        begin
                begin
                  s = Readline.readline("exe> ",true).strip
				  if s == "endme" then break end
                  eval ("puts ' = ' + (#{s}).to_s"),bind
                rescue => e
                  puts " > Error ocurred: #{e.backtrace[0]}: #{e.message}"
                end while true
        ensure
                nil
        end
end

$ips = []
doc = Nokogiri::XML(File.open("201709240539 Intense scan on 172.20.164.1_24.xml"))
doc.search('//nmaprun/host').each do |host|
    
    ip = host.at("address").attr("addr")
	ports = []
	host.at("ports").elements.each{|port| 
	    next if port.node_name!= "port"
		portNum = port.attr("portid")
		protocol = port.attr("protocol")
		state = port.at("state").attr("state")
		serviceName = port.at("service").attr("name")
		serviceName = "-" if serviceName.nil?
		serviceProduct = port.at("service").attr("product")
		serviceProduct = "-" if serviceProduct.nil?
		if state == "open"
			ports.push({:portNum => portNum, :protocol => protocol, :serviceName => serviceName,  :serviceProduct => serviceProduct})
		end
	}
	$ips.push({:ip => ip,:ports => ports})
end

def toword a,file_path
    result = ''
	mrdIP = a[:ip]
	row_ports = a[:ports].map{|b|  TABLEROW_PORTS.gsub("MRDPORT",CGI.escapeHTML(b[:portNum])).gsub("MRDNAME",CGI.escapeHTML(b[:serviceName])).gsub("MRDPROD",CGI.escapeHTML(b[:serviceProduct]))  }.join
	
	Zip::File.open(file_path) do |zipfile|
	  files = zipfile.select(&:file?)
	  files.each do |zip_entry|
		if(zip_entry.name == "word/document.xml")
			s = zipfile.read(zip_entry.name).force_encoding("utf-8")
			s.gsub!("MRDIP",CGI.escapeHTML(mrdIP) )
			s.gsub!(TABLEROW_PORTS,row_ports)
				
				
			start__ = s.index(BODY_START) + BODY_START.size
			end__ = s.index(BODY_END)
			
			#x binding
			
			result =s[start__...end__]
		end
	  end
	  zipfile.commit
	end
	return result
end

def export2word
    file_path = "./temp - Ports.docx"
    cntt = ''
	$ips.each{|a| 
	    if a[:ports].size == 0 then next end
		cntt += toword a, file_path
	}	
	zip_file_name = "./ports.docx"
	FileUtils.cp file_path,zip_file_name
	Zip::File.open(zip_file_name) do |zipfile|
	  files = zipfile.select(&:file?)
	  files.each do |zip_entry|
		if(zip_entry.name == "word/document.xml")
			s = zipfile.read(zip_entry.name).force_encoding("utf-8")
			start__ = s.index(BODY_START) + BODY_START.size
			end__ = s.index(BODY_END)
			# cntt
			s2 = s[0...start__] +  cntt  + s[end__..-1]
			zipfile.get_output_stream(zip_entry.name){ |f| f.puts s2 }
		end
	  end
	  zipfile.commit
	end
	
end


x binding