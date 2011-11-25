"use strict";
// Console interface to the Nano-VM.

function println(s)
{
	WSH.echo(s);
}

function vm_set_port_callback(value) {
}

function vm_get_port_callback() {
}

var regnames = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
function vm_set_register_callback(regno, value) {
    println(regnames.charAt(regno) + " << " + value);
}

function read_text_file(filename) {
    var forReading = 1;

    // define array to store lines. 
    var contents = ""

    // First, we need a File System!
    var fs = new ActiveXObject("Scripting.FileSystemObject");
    var f = fs.GetFile(filename);

    // Open the file for reading, read it all and close the file.
    var is = f.OpenAsTextStream(forReading, 0);
    while (!is.AtEndOfStream) {
        contents += is.ReadLine() + "\n";
    }
    is.Close();

    return contents;
}

function process_file(filename) {
    println("Reading input file: " + filename);
    var contents = read_text_file(filename);
    println("Contents: " + contents);

    // Load&Execute!
    nanovm_load_code(contents);
    while (nanovm_step() > 0) {
        // pass the time!
    }
}

function main(args) {
    nanovm_init(println, vm_set_port_callback, vm_get_port_callback, vm_set_register_callback);

    for (i = 0; i < args.length; ++i) {
        process_file(args(i));
    }
    return 0;
}

main(WScript.Arguments);

