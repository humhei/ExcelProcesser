#if FAKE
#r "paket:
source https://api.nuget.org/v3/index.json
nuget FPublisher
nuget Fake.Core.Target //"
#endif
#if !FAKE
#r "Facades/netstandard"
#r "netstandard"
#endif
#load "./.fake/build.fsx/intellisense.fsx"
open Fake.Core
open FPublisher.Roles
open FPublisher
open FPublisher.Nuget
open FPublisher.Git


let buildServer =
    BuildServer.create
        { BuildServer.Config.DefaultValue
            with
                LocalNugetServer = Some NugetServer.DefaultBaGetLocal
                LoggerLevel = Logger.Level.Normal }


Target.create "Workspace.CleanBinAndObj" <| fun _ ->
    Workspace.cleanBinAndObj buildServer.Collaborator.Workspace

Target.create "Workspace.CreateDefaultSln" <| fun _ ->
    Workspace.createDefaultSln false buildServer.Collaborator.Workspace

Target.create "NonGit.Build" <| fun _ ->
    BuildServer.run (!^ (NonGit.Msg.Build None))  buildServer

Target.create "NonGit.Test" <| fun _ ->
    BuildServer.run (!^ NonGit.Msg.Test) buildServer

Target.create "Forker.PublishToLocalNugetServer" <| fun _ ->
    BuildServer.run (!^ (Forker.Msg.PublishToLocalNugetServer)) buildServer


Target.create "Collaborator.NextRelease" <| fun _ ->
    BuildServer.run (!^ Collaborator.Msg.NextRelease) buildServer

Target.create "RunCI" <| fun _ ->
    BuildServer.run (BuildServer.Msg.RunCI) buildServer

Target.create "Default" ignore

"Default"
    <= "NonGit.Test"

Target.runOrDefault "Default"