require 'win32ole'
require 'icalendar'

class Attendee
  attr :mailto, true
  attr :params, true

  def initialize(mailto, params={})
    self.mailto = mailto
    self.params = params
  end

  def property_name
    param_str = ""
    params.each do |key, value|
    	param_str << ";" if param_str.empty?
    	param_str << "#{key}=#{value}"
    end
    "ATTENDEE#{param_str}"
  end

  def value
    "MAILTO:#{mailto}"
  end
end

groupwise = WIN32OLE.new("NovellGroupWareSession")
# 2nd arg is not password, so leave blank
account = groupwise.Login("your-username", "")
path_to_host = account.PathToHost
cal = Icalendar::Calendar.new
cal.custom_property("METHOD", "PUBLISH")
cal.custom_property("X-WR-CALNAME;VALUE=TEXT", "ASRS GroupWise Calendar")
message_list = []
account.Calendar.Messages.each do |message|
  next unless message.ClassName == "GW.MESSAGE.APPOINTMENT"
  next if message_list.include? message.CommonMessageID
  message_list << message.CommonMessageID
  event = cal.event  # This automatically adds the event to the calendar
  the_date = ParseDate.parsedate(message.CreationDate)
  event.timestamp = DateTime.civil(*the_date.compact!)
  event.location = message.Place
  event.organizer = "MAILTO:#{message.Sender.EmailAddress}"
  event.property_params["ORGANIZER"] = {}
  event.property_params["ORGANIZER"]["CN"] = message.FromText
  message.ExpandedRecipients.each do |recipient|
  	if !recipient.Address.nil?
  		address = recipient.Address.EmailAddress
  	else
  		address = recipient.EmailAddress
  	end
  	attendee = Attendee.new(address, {"CN" => recipient.DisplayName})
  	event.custom_property attendee.property_name, attendee.value
  end
  event.summary = message.subject.PlainText
  event.description = message.BodyText.PlainText
  the_date = ParseDate.parsedate(message.StartDate)
  event.start = DateTime.civil(*the_date.compact!)
  the_date = ParseDate.parsedate(message.EndDate)
  event.end = DateTime.civil(*the_date.compact!)
end

cal.to_ical.each_line do |line|
	puts line.chomp
end
