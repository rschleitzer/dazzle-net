using Dazzle.Cli;
using System.Reflection;

// Add bundled SGML catalog to args
var assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;
var catalogPath = Path.Combine(assemblyDir, "sgml", "catalog");

var newArgs = new List<string> { "-c", catalogPath };
newArgs.AddRange(args);

var app = new DazzleApp();
return app.run(newArgs.ToArray());
