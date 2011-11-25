"use strict";

var _anttivm_println_callback = function (line) { };
var _anttivm_set_port_callback = function (value) { };
var _anttivm_get_port_callback = function () { };
var _anttivm_set_register_callback = function (regno, value) { };

// Initialize the VM system. Call it once on start-up.
// println_callback: (s : string) -> unit
// set_port_callback: (port_value : integer) -> unit
// get_port_callback: () -> integer
// set_register_callback: (register:char) -> (value:integer) -> unit
function anttivm_init(println_callback, set_port_callback, get_port_callback, set_register_callback) {
    _anttivm_println_callback = println_callback;
    _anttivm_set_port_callback = set_port_callback;
    _anttivm_get_port_callback = get_port_callback;
    _anttivm_set_register_callback = set_register_callback;
}

var _anttivm_code = "";   // string
var _anttivm_ip = 1;         // integer. EOF, initially.

var _anttivm_stack = [];

var _anttivm_registers = [];

// Load some code, but do not execute it.
function anttivm_load_code(code) {
    _anttivm_code = code;
    _anttivm_ip = 0;
    _anttivm_stack = [];
    _anttivm_registers = [];
}

// Return instruction pointer.
function anttivm_get_ip() {
    return _anttivm_ip >= _anttivm_code.length ? "EOF" : _anttivm_ip;
}

function anttivm_get_instruction() {
    if (_anttivm_ip<0 || _anttivm_ip>=_anttivm_code.length) {
        return "N/A";
    }
    else {
        var opcode = _anttivm_code.charCodeAt(_anttivm_ip);
        if (opcode >= 0x30 && opcode <= 0x39) {
            // Digits.
            var number = opcode - 0x30;
            for (var ip = _anttivm_ip + 1; ip < _anttivm_code.length; ++ip) {
                var number2 = _anttivm_code.charCodeAt(ip) - 0x30;
                if (number2 >= 0 && number2 <= 9) {
                    number = number * 10 + number2;
                }
                else {
                    break;
                }
            }
            return number;
        }
        else {
            return _anttivm_code.charAt(_anttivm_ip);
        }
    }
}

// Return copy of the stack.
function anttivm_get_stack() {
    return _anttivm_stack.slice(0);
}

// Return copy of the registers.
function anttivm_get_registers() {
    return _anttivm_registers.slice(0);
}

// Execute one step of the code.
function anttivm_step() {
    if (_anttivm_ip < _anttivm_code.length) {
        var opcode = _anttivm_code.charCodeAt(_anttivm_ip);
        ++_anttivm_ip;

        if (opcode >= 0x61 && opcode <= 0x7A) {
            // Small letters.
            _anttivm_stack.push(_anttivm_registers[opcode - 0x61]);
        }
        else if (opcode >= 0x41 && opcode <= 0x5A) {
            // Capital letters.
            var x = _anttivm_stack.pop();
            var regno = opcode - 0x41;
            _anttivm_registers[regno] = x;
            _anttivm_set_register_callback(regno, x);
        }
        else if (opcode >= 0x30 && opcode <= 0x39) {
            // Digits.
            var number = opcode - 0x30;
            for (; _anttivm_ip < _anttivm_code.length; ++_anttivm_ip) {
                var number2 = _anttivm_code.charCodeAt(_anttivm_ip) - 0x30;
                if (number2 >= 0 && number2 <= 9) {
                    number = number * 10 + number2;
                }
                else {
                    break;
                }
            }
            _anttivm_stack.push(number);
        }
        else {
            switch (opcode) {
                case 0x0A: // CR
                    break;
                case 0x0D: // LF
                    break;
                case 0x20: // ' '
                    break;
                case 0x2A: // '*'
                    var arg2 = _anttivm_stack.pop();
                    var arg1 = _anttivm_stack.pop();
                    _anttivm_stack.push(arg1 * arg2);
                    break;
                case 0x2B: // '+'
                    var arg2 = _anttivm_stack.pop();
                    var arg1 = _anttivm_stack.pop();
                    _anttivm_stack.push(arg1 + arg2);
                    break;
                case 0x2D: // '-'
                    var arg2 = _anttivm_stack.pop();
                    var arg1 = _anttivm_stack.pop();
                    _anttivm_stack.push(arg1 - arg2);
                    break;
                case 0x2F: // '/'
                    var arg2 = _anttivm_stack.pop();
                    var arg1 = _anttivm_stack.pop();
                    _anttivm_stack.push(arg1 / arg2);
                    break;
                default:
                    _anttivm_println_callback('Unknown opcode: ' + opcode);
                    break;
            }
        }
        return 1;
    }
    else {
        return 0;
    }
}
