require 'date'
require 'net/http'
require 'json'
require 'yaml'
require 'optparse'
require 'optparse/date'

class Timesheet

  def self.from_file(filename, client_id)
    timehash = YAML::load_file(filename).to_hash
    Timesheet.new(clinet_id: client_id, timehash: timehash)
  end


  def initialize(args)
    @client_id = args[:client_id]
    @timehash =  args[:timehash] || { 
      "Action" => "Edit", 
      "Type" => "Replicon.TimeSheet.Domain.Timesheet", 
      "Operations" => [ 
        {
          "__operation" => "CollectionClear", "Collection" => "TimeRows" 
        }
      ]
    }

    @timehash["Identity"] = args[:timesheet_id] unless args[:timehash]
  end

  def enter_time(task_id, date, duration)
    rows = @timehash["Operations"]
    row = rows[1..-1].find do |row|
      row["Operations"][0]["Task"]["Identity"] == task_id
    end

    if !row
      row =       {
        "__operation" => "CollectionAdd", 
        "Collection" => "TimeRows", 
        "Operations" =>  [ 
          {
            "__operation" => "SetProperties", 
            "Task" => { "__type" => "Replicon.Project.Domain.Task", "Identity" => task_id },
            "Client" => { "__type" => "Replicon.Project.Domain.Client", "Identity" => @client_id }
          }
        ]
      }

      rows << row
    end

    row["Operations"] <<  {
      "__operation" => "CollectionAdd", 
      "Collection" => "Cells", 
      "Operations" => [
        {
          "__operation" => "SetProperties", 
          "CalculationModeObject" => {
            "__type" => "Replicon.TimeSheet.Domain.CalculationModeObject", "Identity" => "CalculateInOutTime" 
          }
        },
        {
          "__operation" => "SetProperties", 
          "EntryDate" => { "__type" => "Date", "Year" => date.year, "Month" => date.month, "Day" => date.day }, 
          "Duration" => { "__type" => "Timespan", "Hours" => duration }, 
        },
      ] 
    }
  end

  def to_hash
    @timehash
  end
end

class Replicon
  URL = "https://na1.replicon.com/orbitzllc/remoteAPI/remoteapi.ashx/8.29.28/"
  URI = URI(URL)

  HEADERS = {
    "X-Replicon-Security-Context" => "User",
    "Content-Type" => "application/json"
  }

  attr_reader :current_timesheet, :client

  def initialize(clientName, userId, password, verbose = false)
    raise "Client name must be supplied" unless clientName
    @verbose = verbose
    @clientName = clientName
    @userId = userId
    @password = password
    @client = lookup_client(@clientName)
    @user = user(@userId)
    @current_timesheet = timesheet(Date.today)
  end
  
  def beginSession
    {
      "Action" => "BeginSession"   
    }
  end

  def endSession 
    {
      "Action" => "EndSession"   
    }
  end
  
  def submitTimesheet(timesheet = @current_timesheet)
    {
      "Action" => "Edit",
      "Type" => "Replicon.TimeSheet.Domain.Timesheet",
      "Identity" => timesheet,
      "Operations" => [ {"__operation" => "Submit" } ]
      }
  end
  
  def lookup_client(name)
    queryForIdentity({ "Action" => "Query", 
      "QueryType" => "ClientByName", 
      "DomainType" => "Replicon.Domain.Client", 
      "Args" => [ name ] 
    })
  end

  def user(wwid = @wwid)
    userByLoginName = { 
      "Action" => "Query", 
      "QueryType" => "UserByLoginName", 
      "DomainType" => "Replicon.Domain.User",
      "Args"=> [ wwid ]
    }

    queryForIdentity(userByLoginName)
  end

  def timesheet(date)
    timesheetByUserDate = { 
      "Action" => "Query", 
      "QueryType" => "TimesheetByUserDate", 
      "DomainType" => "Replicon.TimeSheet.Domain.Timesheet", 
      "Args" => [ 
        { "__type" => "Replicon.Domain.User", "Identity" => @user }, 
        { "__type" => "Date", "Year" => date.year, "Month" => date.month, "Day" => date.day } ], 
      "Load" => [ 
        {
          "Relationship" => "TimeRows", 
          "Load" => [ 
            { "Relationship" => "Activity" }, 
            { "Relationship" => "Cells" } ] 
        } 
      ]
    }

    Timesheet.new(client_id: @client, timesheet_id: queryForIdentity(timesheetByUserDate))
  end

  def project(projectCode)
    projectByCode = { 
      "Action" => "Query", 
      "QueryType" => "ProjectByCode", 
      "DomainType" => "Replicon.Project.Domain.Project", 
      "Args" => [ projectCode ]
    }
    queryForIdentity(projectByCode)
  end

  def tasks(taskName, project)
    appOpenTasksByUserAndProject = { 
      "Action" => "Query", 
      "QueryType" => "AllOpenTasksByUserAndProject", 
      "DomainType" => "Replicon.Project.Domain.Task", 
      "Args" => [ 
        { "__type" => "Replicon.Domain.User", "Identity" => @user },
        { "__type" => "Replicon.Project.Domain.Project", "Identity" => project } ] 
      }
    execute(appOpenTasksByUserAndProject)["Value"][1..-1].find { |task| task["Properties"]["Name"] == taskName  }["Identity"]
  end

  def queryForIdentity(queryAction)
    execute(queryAction)["Value"][1]["Identity"]
  end

  def execute(*actions)
    req = Net::HTTP::Post.new(URI, initheader = HEADERS)
    req.basic_auth @userId, @password
    req.body = [ beginSession, actions, endSession ].flatten(1).to_json

    puts "Request:
            #{JSON.pretty_generate(JSON.parse(req.body))}" if @verbose


    res = Net::HTTP.start(URI.hostname, URI.port, :use_ssl => URI.scheme == 'https') do |http|
      http.request(req)
    end

    response = JSON.parse(res.body)
    puts "Response #{res.code} #{res.message}:
            #{JSON.pretty_generate(response)}" if @verbose

    raise "#{response['Message']}:
              Request: #{JSON.pretty_generate(JSON.parse(req.body))}
              Response (#{res.code}): #{JSON.pretty_generate(response)}" if response["Status"] == "Exception" || res.code != "200"

    response
  end
end

def self.parse(args)
  creds = YAML::load_file(File.join(ENV['HOME'], '.repliconrc'))

  options = OpenStruct.new
  options.verbose = false
  options.submit = false
  options.username = creds['username']
  options.password = creds['password']
  options.client_name = creds['client']
  options.project_number = creds['default_project']
  options.task_name = creds['default_task']
  options.dates = []
  options.hours = 8
  options.save = false

  opt_parser = OptionParser.new do |opts|
    opts.banner = "Usage: replicon.rb [options]"

    opts.on("--username USERNAME", String, "Replicon username, default is defined in ~/.repliconrc") do |username|
      options.username = username
    end

    opts.on("--password PASSWORD", "Replicon password, default is defined in ~/.repliconrc") do |password|
      options.password = password
    end

    opts.on("--client CLIENT", "Replicon client name, default is defined in ~/.repliconrc") do |client|
      options.client = client
    end

    opts.on("-p", "--project PROJECT", "Replicon project number, default is defined in ~/.repliconrc") do |project_number|
      options.project_number = project_number
    end

    opts.on("-t", "--task TASK", "Replicon task name, default is defined in ~/.repliconrc") do |task_name|
      options.task_name = task_name
    end

    opts.on("-d", "--date DATE", Date, "The dates to apply the hours to. Multiple allowed") do |date|
      options.dates << date
    end

    opts.on("-h", "--hours HOURS", Integer, "The number of hours to be entered on the date for the project task.") do |hours|
      options.hours << hours
    end

    opts.on("--save [FILE]", "Save the timesheet to a file. File name can be provided.") do |file|
      options.save = true
      options.save_file = file || nil
    end

    opts.on("--load [FILE]", "Load the timesheet from a file. File name can be provided.") do |file|
      options.load = true
      options.load_file = file || nil
    end

    opts.on("--file FILE", "File name to be used for loading and saving.") do |file|
      options.file = file
    end

    opts.on("-s", "--submit", "Submit the timesheet.") do |submit|
      options.submit = submit
    end

    opts.on("-v", "--[no-]verbose", "Run verbosely") do |v|
      options.verbose = v
    end

    opts.on_tail("-h", "--help", "Show this message") do
      puts opts
      exit
    end
  end

  opt_parser.parse!(args)
  options.dates << Date.today if options.dates.empty?
  options.file ||= "WK#{options.dates[0].strftime("%U")}_timesheet.yml"
  options.save_file ||= options.file 
  options.load_file ||= options.file

  options

end

options = parse(ARGV)
puts options if options.verbose

replicon = Replicon.new(options.client_name, options.username, options.password, options.verbose)
project_id = replicon.project(options.project_number)
task_id = replicon.tasks(options.task_name, project_id)

begin
  timesheet = Timesheet.from_file(options.load_file, replicon.client) 
rescue Exception => e
  puts "Unable to load from #{options.load_file}"
end if options.load
timesheet ||= replicon.timesheet options.dates[0]

options.dates.each do |date| 
  timesheet.enter_time(task_id, date, options.hours) 
end

File.open(options.save_file, 'w') { |fo| fo.puts timesheet.to_yaml } if options.save

timehash = timesheet.to_hash

if options.submit
  replicon.execute(timehash, replicon.submitTimesheet(timehash["Identity"]))
else
  replicon.execute(timehash)
end