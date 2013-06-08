# CV3 / Quickbooks integration script
# Version 1.0b1
# copyright 2010 CommerceV3, Inc
# Blake Ellis <blake@commercev3.com>
# 6/2/2010

# Configuration - the following variables need to be set per CV3 storefront
@cv3_user = "store_name"         			# cv3 user name
@cv3_password = "store_password"				# cv3 password
@cv3_api_key = "api_key"					# cv3 API key
@qb_service_name = "cv3 web wholesale"			# the name of the programming accessing Quickbooks
                         					# NOTE: the first time this script connects to Quickbooks,
											# Quickbooks will present an authorization dialog showing
											# that a program by this name is attempting access. This 
											# dialog must be OK'd manually by the QB admin. Subsequent
											# requests will not pop the dialog. If you change this name
											# you will have to approve a new authorization dialog.

@include_source_code = "Wholesale"			# if you'd like to ONLY process orders with a certain CV3
											# source code, put it here. This is useful if you're running
											# two instances of this script to put orders into two different
											# quickbooks files. Best example would be seperating "retail"
											# and "wholesale" orders
											
@pass_cc = false							# 0 = don't pass credit card info, 1 = pass cc info to Quickbooks

@qb_unknown_item = "Unknown"				# Quickbooks requires a match per line item on Sales Orders, so
											# you need to create a QB Inventory Part named exactly this for
											# any line items where the SKU doesn't match. This script will put
											# the SKU and product name in the Sales Order line item description
											# so you can match it manually after the sales order is created.
											
@qb_tax_service_name = "Sales Tax"			# There must be Tax and Shipping "items" in Quickbooks that are
@qb_shipping_service_name = "Shipping"		# set up as "Service" type. Put the specific item names here.

@qb_gift_message_service_name = "Gift Message"	#As above, a Service Item must be set up in QB so that we
												# can print a Gift Message on the Order, this is the item name

#@qb_modified_from = Time.now - 2.hours		# These variables set the time period this script searchs QB for
#@qb_modified_to = Time.now + 2.hours		# orders to match for confirmation before removing pending orders
											# from CV3. Currently it's set for two hours into the past and 
											# future, to accommodate differences in computer times. For high
											# order volume sites this might need to be throttled down.
											
@remove_from_pending = true				# true will confirm orders are in QB and then remove from CV3 pending

# Advanced Configuration - these variables manage cv3 API calls
@cv3_wsdl_address = "https://service.commercev3.com/index.php?wsdl"
@cv3_authenticate_xml = "<authenticate><user>#{@cv3_user}</user><pass>#{@cv3_password}</pass><serviceID>#{@cv3_api_key}</serviceID></authenticate>"
@cv3_order_request_xml = "<reqOrders><reqOrderNew/></reqOrders>"

# ---------------- do not edit anything below this line! ------------------- #
require 'rubygems'
require 'quickbooks'  #access Quickbooks API, purchase at http://behindlogic.com/
require 'savon'       #connect to SOAP-based web services
require 'base64'      #encode & decode base64 text
require 'crack'       #XML parser, part of HTTParty

puts "-----------------------------------------------"
puts "CommerceV3 / Quickbooks Integration version 1.0"
puts "       copyright 2010 CommerceV3, Inc."
puts "-----------------------------------------------"
puts ""
sleep 2

# turn of logging... comment this out to see SOAP traffic
Savon::Request.log = false

# initialize Savon with CV3's WSDL address
client = Savon::Client.new @cv3_wsdl_address

# send SOAP request and store response
puts "Contacting CommerceV3 API..."
response = client.cv3_data do |soap|
  soap.input = "CV3Data"
  soap.action = "CV3Data"
  soap.body = Base64.encode64("<CV3Data version=\"2.0\"><request>#{@cv3_authenticate_xml}<requests>#{@cv3_order_request_xml}</requests></request></CV3Data>")
end

# decode and parse CV3's SOAP response
data = response.to_hash
data.each do |key,val|
  val.each do |key2, val2|
	if key2.to_s == "return"
	  final = Base64.decode64(val2.to_s)
	  @cv3_data = Crack::XML.parse(final)
	end
  end
end

# Connect to Quickbooks
# currently this connection requires the specified QB file to be open on the same computer
# running this script.

Quickbooks.use_adapter :ole, @qb_service_name
#Quickbooks.use_adapter :test
#Quickbooks.use_adapter :ole, {:application_name => "cv3", :file => "sav bee 2009", :username => "Admin", :password => "HoneyBee1" }

# if there are no orders, notify the user and abort the script.
unless @cv3_data["CV3Data"]["orders"]
  puts ""
  puts "copyright 2010 CommerceV3, Inc"
  puts "contact Blake Ellis <blake@commercev3.com> for more information"
  puts ""
  print "no orders to process, cleaning up ***"
  (1..10).each do |num|
    print "*"
    sleep 1
  end 
 abort("no pending orders")
end

puts ""

# The data structure is different if there is only one order, so here we 
# check and adjust it if necessary.

if @cv3_data["CV3Data"]["orders"]["order"].kind_of? Array
  puts "Found #{@cv3_data["CV3Data"]["orders"]["order"].size} orders..."
else
  puts "Found 1 order..."
  order = @cv3_data["CV3Data"]["orders"]["order"]
  @cv3_data["CV3Data"]["orders"]["order"] = [order]
end

# set a counter for command line display purposes
x = 0

# loop through the pending orders and process them into Quickbooks
@cv3_data["CV3Data"]["orders"]["order"].each do |order|
  x = x+1 	#increment the counter for command line display
  puts ""
  puts "processing order #{x}..."
  # set the Quickbooks Customer Name, which must be unique. For our purposes we use
  # the following convention:
  #
  #  Blake Ellis <blake@commercev3.com>
  #
  # QB has a 41 character limit on this field, so we also crop it if necessary.
  
  @qb_full_name = "#{order["billing"]["lastName"]}, #{order["billing"]["firstName"]} <#{order["billing"]["email"]}>"
  if @qb_full_name.length > 40
    @qb_full_name = @qb_full_name[0,40]
  end

  # if we're excluding a certain source code, we skip the order here
  if order["sourceCode"]!= @include_source_code
  	puts "Order ##{x} - skipping {@include_source_code} order from #{@qb_full_name}..." 
  else
	puts "Order ##{x} - downloading order from #{@qb_full_name}..."
	  
	# check if this customer exists?
	@customer = QB::Customer.first(:FullName => @qb_full_name)
	
	# if we have a match, update the customer record in QB
	if @customer
	    
		puts "customer exists, updating #{@qb_full_name}..."
		
		#This script updates all customer fields in Quickbooks with the latest information they enter
		@customer.Name = @qb_full_name
		@customer.FirstName = order["billing"]["firstName"]
		@customer.LastName = order["billing"]["lastName"]
		@customer.Email = order["billing"]["email"]
		@customer.Phone = order["billing"]["phone"]
	
		@customer.BillAddress = {	:Addr1 => "#{order["billing"]["firstName"]} #{order["billing"]["lastName"]}",
									:Addr2 => order["billing"]["address1"],
									:Addr3 => order["billing"]["address2"],
									:City => order["billing"]["city"],
									:State => order["billing"]["state"],
									:PostalCode => order["billing"]["zip"],
									:Country => order["billing"]["country"]
								}
	    
		# if we're passing credit cards into quickbooks, we do it here.
		if @pass_cc
			if order["payMethod"] == "creditcard" && order["billing"]["CCInfo"]["CCNum"].to_i > 1
				@customer.CreditCardInfo = {:CreditCardNumber => order["billing"]["CCInfo"]["CCNum"],
											:ExpirationMonth => order["billing"]["CCInfo"]["CCExpM"],
											:ExpirationYear => order["billing"]["CCInfo"]["CCExpY"],
											:NameOnCard => order["billing"]["CCInfo"]["CCName"]
											}
			end
		end
		
		@customer.valid?
		@customer.save
	
	  else
	  	# If we don't have a customer, create a new one
	    puts "customer doesn't exist, adding #{@qb_full_name}..."
		
		@customer = QB::Customer.new(	:Name => @qb_full_name, 
										:FirstName => order["billing"]["firstName"], 
										:LastName => order["billing"]["lastName"],
										:Email => order["billing"]["email"],
										:Phone => order["billing"]["phone"],
										:BillAddress => {	:Addr1 => "#{order["billing"]["firstName"]} #{order["billing"]["lastName"]}",
															:Addr2 => order["billing"]["address1"],
															:Addr3 => order["billing"]["address2"],
															:City => order["billing"]["city"],
															:State => order["billing"]["state"],
															:PostalCode => order["billing"]["zip"],
															:Country => order["billing"]["country"]
														}
									)
		if @pass_cc
			if order["payMethod"] == "creditcard"
				@customer.CreditCardInfo = {:CreditCardNumber => order["billing"]["CCInfo"]["CCNum"],
											:ExpirationMonth => order["billing"]["CCInfo"]["CCExpM"],
											:ExpirationYear => order["billing"]["CCInfo"]["CCExpY"],
											:NameOnCard => order["billing"]["CCInfo"]["CCName"]
											}
			end
		end
		
		@customer.valid?
		@customer.save
	  end
	
	  # Now that we have a customer to attach to the Sales Orders, we create a
	  # Sales Order for each CV3 "ship to" address.
	
	  order["shipTos"].each do |shipTo|
		puts "adding a sales order for #{@qb_full_name}..."
	
		# add line items, one per shipToProducts
		shipTo.each do |shipToProducts|
			if shipToProducts["name"]

				@sales_order = QB::SalesOrder.new(	:CustomerRef => {:FullName => @qb_full_name},
													:ClassRef => {:FullName => "Retail"},
													:SalesRepRef => {:FullName => "web"},
													:CustomerMsgRef => {:FullName => "Thank you for your business."},
													:CustomerRef => {:FullName => @qb_full_name},
													:PONumber => order["orderID"],
													:ShipAddress => {	:Addr1 => shipToProducts["name"],
																		:Addr2 => shipToProducts["address1"],
																		:Addr3 => shipToProducts["address2"],
																		:City => shipToProducts["city"],
																		:State => shipToProducts["state"],
																		:PostalCode => shipToProducts["zip"],
																		},
													:Memo => order["comments"],
													:IsToBePrinted => false
													 )
	
	
				shipToProducts.each do |shipToProduct|
					
					shipToProduct.each do |line|
						if line && line["shipToProduct"]
						  if line["shipToProduct"].kind_of?(Array)
						    
							# we have an array of products
							line["shipToProduct"].each do |prod|
								#puts "prod size is string -#{prod.is_s?}-... #{prod}"
								#find the item, if it doesn't exist find our "Unknown" item
								#QB REQUIRES A MATCH HERE OR THE SCRIPT WILL FAIL

								  qb_line_item = QB::ItemInventory.first(:FullName => prod["altID"])
								  unless qb_line_item
								    qb_line_item = QB::ItemInventory.first(:FullName => @qb_unknown_item)
								  end
								  #puts prod.to_yaml
								  if prod["quantity"].to_i > 0
									  if @sales_order[:SalesOrderLine]
										@sales_order[:SalesOrderLine] << {
											:ItemRef => qb_line_item.to_ref,
											:Desc => "#{prod["altID"]} - #{qb_line_item.SalesDesc}",
											:Quantity => prod["quantity"],
											:Rate => prod["price"] }
									  else
										@sales_order[:SalesOrderLine] = [{
											:ItemRef => qb_line_item.to_ref,
											:Desc => "#{prod["altID"]} - #{qb_line_item.SalesDesc}",
											:Quantity => prod["quantity"],
											:Rate => prod["price"] }]
									  end #end if/else
								  end # end if
							  end # end each
							else
								# we have a single product
									qb_line_item = QB::ItemInventory.first(:FullName => line["shipToProduct"]["altID"])
									unless qb_line_item
										qb_line_item = QB::ItemInventory.first(:FullName => @qb_unknown_item)
									end
									if line["shipToProduct"]["quantity"].to_i > 0
										@sales_order[:SalesOrderLine] = [{
											:ItemRef => qb_line_item.to_ref,
											:Desc => "#{line["shipToProduct"]["altID"]} - #{qb_line_item.SalesDesc}",
											:Quantity => line["shipToProduct"]["quantity"],
											:Rate => line["shipToProduct"]["price"] }]
									end
	
							end
						end
					end
				end
				# if there is a gift message, add it as a line item so it prints on the sales order,
				# note there must be a Service Item in the QB Item List for Gift Message
				if shipToProducts["message"]
					qb_line_item_gift_message = QB::ItemService.first(:FullName => @qb_gift_message_service_name)
					@sales_order[:SalesOrderLine] << {
												:ItemRef  => qb_line_item_gift_message.to_ref,
												:Desc => shipToProducts["message"]}
				end
											
				# add line items for tax and shipping
				qb_line_item_tax = QB::ItemService.first(:FullName => @qb_tax_service_name)
				qb_line_item_shipping = QB::ItemService.first(:FullName => @qb_shipping_service_name)
				@sales_order[:SalesOrderLine] << {
											:ItemRef  => qb_line_item_tax.to_ref,
											:Desc => "Sales Tax",
											:Amount => shipToProducts["tax"] }
				
				@sales_order[:SalesOrderLine] << {
											:ItemRef => qb_line_item_shipping.to_ref,
											:Desc => "Ship on #{shipToProducts["shipOn"]} via #{shipToProducts["shipMethodCode"]}",
											:Amount => shipToProducts["shipping"] }
				@sales_order.save
			end
		end
	  end

  end  # end if excluded
end

# We've inserted all orders into Quickbooks, now we loop through and confirm each order to prepare to
# remove it from CV3's pending orders section.
puts "Confirming Orders in QB..."

# this array holds all confirmed sales orders in preperation for removal from CV3 pending orders
@confirmed_order_array = Array.new

# pull recent orders from QB
todayQBO = QB::SalesOrder.all(:ModifiedDateRangeFilter => {:FromModifiedDate => Time.now-2.hours, :ToModifiedDate => Time.now+2.hours})
#puts "checking #{todayQBO.size} recent orders in QB"

@cv3_data["CV3Data"]["orders"]["order"].each do |order|

  if order["sourceCode"]== @include_source_code
	  todayQBO.each do |today_order|
		#puts "testing #{today_order[:RefNumber]}( PO# #{today_order[:PONumber]}) against #{order["orderID"]}"
		if today_order["PONumber"].to_i == order["orderID"].to_i
			puts "Order #{order["orderID"]} confirmed..."
			#add order ID to confirmed order array
			@confirmed_order_array << order["orderID"]
		end
	  end
  end
end

# go back to CV3 and remove all confirmed orders from the pending order section
if @remove_from_pending
	if @confirmed_order_array.size > 0
		puts "removing confirmed orders from CV3..."
		conf_response_string = '<CV3Data version="2.0">'
		conf_response_string << "<request>#{@cv3_authenticate_xml}</request><confirm><orderConfirm>"
		@confirmed_order_array.each do |co|
			conf_response_string << "<orderConf>#{co}</orderConf>"
		end
		conf_response_string << "</orderConfirm></confirm></CV3Data>"
	
		# send delete back to the API to remove the order from CV3
		confirm_response = client.cv3_data do |soap|
		  soap.input = "CV3Data"
		  soap.action = "CV3Data"
		  soap.body = Base64.encode64(conf_response_string)
		end
	
		confirm_data = confirm_response.to_hash
		confirm_data.each do |key,val|
		  val.each do |key2, val2|
			if key2.to_s == "return"
			  final = Base64.decode64(val2.to_s)
			  @confirm_cv3_data = Crack::XML.parse(final)
			end
		  end
		end
	end
else
	puts "Orders not removed from CV3 pending per configuration"
end

puts ""
puts "copyright 2010 CommerceV3, Inc"
puts "contact Blake Ellis <blake@commercev3.com> for more information"
puts ""
sleep 2
print "cleaning up ***"
(1..10).each do |num|
  print "*"
  sleep 1
end 