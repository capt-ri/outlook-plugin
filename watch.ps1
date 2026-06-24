# Dev server with hot reload for the Outlook add-in.
# Auto-rebuilds and reloads the browser on any .razor / .cs / .css / .js change.
# Serves https://localhost:7010  (same URL the manifest points at).
$env:DOTNET_WATCH_RESTART_ON_RUDE_EDIT = "true"   # restart instead of prompting on structural edits
dotnet watch run --launch-profile https
