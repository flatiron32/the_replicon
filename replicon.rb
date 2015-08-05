require 'date'
require 'net/http'
require 'json'
require 'yaml'
creds = YAML::load_file(File.join(ENV['HOME'], '.repliconrc'))

class Timesheet


  def initialize(timesheet_id, client_id)
    @client_id = client_id
    @timehash =  { 
      "Action" => "Edit", 
      "Type" => "Replicon.TimeSheet.Domain.Timesheet", 
      "Operations" => [ 
        {
          "__operation" => "CollectionClear", "Collection" => "TimeRows" 
        }
      ]
    }
    @timehash["Identity"] = timesheet_id
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

  def initialize(clientName, userId, password, verbose = false)
    @verbose = verbose
    @clientName = clientName
    @userId = userId
    @password = password
    @client = client(@clientName)
    @user = user(@userId)
    @currentTimesheet = timesheet(Date.today)
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
  
  def submitTimesheet(timesheet = @currentTimesheet)
    {
      "Action" => "Edit",
      "Type" => "Replicon.TimeSheet.Domain.Timesheet",
      "Identity" => timesheet,
      "Operations" => [ {"__operation" => "Submit" } ]
      }
  end
  
  def client(name)
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

    Timesheet.new(queryForIdentity(timesheetByUserDate), @client)
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
              Response: #{JSON.pretty_generate(response)}" if response["Status"] == "Exception"

    response
  end
end

# TODO Scriptify

userId = creds['username']
password = creds['password']
clientName = "OWW"
replicon = Replicon.new(clientName, userId, password)
project_id = replicon.project("101544")
cd_project_id = replicon.project("204601")
task_id = replicon.tasks("Initiate/Expense", project_id)
cd_task_id = replicon.tasks("Develop", cd_project_id)
dates = (Date.new(2015,8,3)..Date.new(2015,8,7)).first 5
timesheet = replicon.timesheet(dates[0])
dates.each do |date| 
  timesheet.enter_time(task_id, date, 4) 
  timesheet.enter_time(cd_task_id, date, 4) 
end
replicon.execute(timesheet.to_hash)
