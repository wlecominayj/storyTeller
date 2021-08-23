// DSL for Project Planning
// Copyright (C) Dmitri Nesteruk, dmitri@activemesa.com, 2009
// All rights reserved.

open System;
open Microsoft.Office.Interop.MSProject;

type Task() =
  [<DefaultValue>] val mutable Name : string
  [<DefaultValue>] val mutable Duration : string
  
type Resource() =
  [<DefaultValue>] val mutable Name : string
  [<DefaultValue>] val mutable Position : string
  [<DefaultValue>] val mutable Rate : int
  
type Group() =
  [<DefaultValue>] val mutable Name : string
  [<DefaultValue>] val mutable Person : Resource
  [<DefaultValue>] val mutable Tasks : Task list
 
type Project() =
  [<DefaultValue>] val mutable Name : string
  [<DefaultValue>] val mutable Resources : Resource list
  [<DefaultValue>] val mutable StartDate : DateTime
  [<DefaultValue>] val mutable Groups : Group list
  
let mutable my_project = new Project()

// dsl constructs
let project name startskey start =
  my_project <- new Project()
  my_project.Name <- name
  my_project.Resources <- []
  my_project.Groups <- []
  my_project.StartDate <- DateTime.Parse(start)
  
let with_rate = -1
let starts = -1
let isa = -1
let done_by = -1
let takes = -1

let hours = 1
let hour = 1
let days = 2
let day = 2
let weeks = 3
let week = 3
let months = 4
let month = 4
  
let resource name isakey position ratekey rate =
  let r = new Resource()
  r.Name <- name
  r.Position <- position
  r.Rate <- rate
  my_project.Resources <- r :: my_project.Resources
  
let group name donebytoken resource =
  let g = new Group()
  g.Name <- name
  g.Person <- my_project.Resources |> List.find(fun f -> f.Name = resource)
  my_project.Groups <- g :: my_project.Groups
  
let task name takestoken count timeunit =
  let t = new Task()
  t.Name <- name
  let dummy = 1 + count
  match timeunit with
  | 1 -> t.Duration <- String.Format("{0}h", count)
  | 2 -> t.Duration <- String.Format("{0}d", count)
  | 3 -> t.Duration <- String.Format("{0}wk", count)
  | 4 -> t.Duration <- String.Format("{0}mon", count)
  | _ -> raise(ArgumentException("only spans of hour(s), day(s), week(s) and month(s) are supported"))
  let g = List.hd my_project.Groups
  g.Tasks <- t :: g.Tasks
  
let prepare (proj:Project) =
  let app = new ApplicationClass()
  app.Visible <- true
  let p = app.Projects.Add()
  p.Name <- proj.Name
  proj.Resources |> List.iter(fun r ->
    let r' = p.Resources.Add()
    r'.Name <- r.Position // position, not name :)
    let tables = r'.CostRateTables
    let table = tables.[1]
    table.PayRates.[1].StandardRate <- r.Rate
    table.PayRates.[1].OvertimeRate <- (r.Rate + (r.Rate >>> 1)))
  // make root task with project name
  let root = p.Tasks.Add()
  root.Name <- proj.Name
  // add groups
  proj.Groups |> List.rev |> List.iter(fun g -> 
    let t = p.Tasks.Add()
    t.Name <- g.Name
    t.OutlineLevel <- 2s
    // who is responsible for this group?
    t.ResourceNames <- g.Person.Position
    // add tasks
    let tasksInOrder = g.Tasks |> List.rev
    tasksInOrder |> List.iter(fun t' ->
        let t'' = p.Tasks.Add(t'.Name)
        t''.Duration <- t'.Duration
        t''.OutlineLevel <- 3s
        // make task follow previous
        let idx = tasksInOrder |> List.findIndex(fun f -> f.Equals(t'))
        if (idx > 0) then 
          t''.Predecessors <- Convert.ToString(t''.Index - 1)
      )
    )
  
// usage
project "F# DSL Article" starts "01/01/2009"
resource "Dmitri" isa "Writer" with_rate 140
resource "Computer" isa "Dumb Machine" with_rate 0

group "DSL Popularization" done_by "Dmitri"
task "Create basic estimation DSL" takes 1 day
task "Write article" takes 1 day
task "Post on CP and wait for comments" takes 1 week

group "Infrastructure Support" done_by "Computer"
task "Provide VS2010 and MS Project" takes 1 day
task "Download and deploy TypograFix" takes 1 day
task "Sit idly while owner waits for comments" takes 1 week

prepare my_project