"use strict";

var _nanovm_println_callback = function (line) { };
var _nanovm_set_port_callback = function (value) { };
var _nanovm_get_port_callback = function () { };
var _nanovm_set_register_callback = function (regno, value) { };

// Initialize the VM system. Call it once on start-up.
// println_callback: (s : string) -> unit
// set_port_callback: (port_value : integer) -> unit
// get_port_callback: () -> integer
// set_register_callback: (register:char) -> (value:integer) -> unit
function nanovm_init(println_callback, set_port_callback, get_port_callback, set_register_callback) {
    _nanovm_println_callback = println_callback;
    _nanovm_set_port_callback = set_port_callback;
    _nanovm_get_port_callback = get_port_callback;
    _nanovm_set_register_callback = set_register_callback;
}

var _nanovm_code = "";   // string
var _nanovm_ip = 1;         // integer. EOF, initially.

var _nanovm_stack = [];

var _nanovm_registers = [];

// Load some code, but do not execute it.
function nanovm_load_code(code) {
    _nanovm_code = code;
    _nanovm_ip = 0;
    _nanovm_stack = [];
    _nanovm_registers = [];
}

// Return instruction pointer.
function nanovm_get_ip() {
    return _nanovm_ip >= _nanovm_code.length ? "EOF" : _nanovm_ip;
}

function nanovm_get_instruction() {
    if (_nanovm_ip<0 || _nanovm_ip>=_nanovm_code.length) {
        return "N/A";
    }
    else {
        var opcode = _nanovm_code.charCodeAt(_nanovm_ip);
        if (opcode >= 0x30 && opcode <= 0x39) {
            // Digits.
            var number = opcode - 0x30;
            for (var ip = _nanovm_ip + 1; ip < _nanovm_code.length; ++ip) {
                var number2 = _nanovm_code.charCodeAt(ip) - 0x30;
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
            return _nanovm_code.charAt(_nanovm_ip);
        }
    }
}

// Return copy of the stack.
function nanovm_get_stack() {
    return _nanovm_stack.slice(0);
}

// Return copy of the registers.
function nanovm_get_registers() {
    return _nanovm_registers.slice(0);
}

// Execute one step of the code.
function nanovm_step() {
    if (_nanovm_ip < _nanovm_code.length) {
        var opcode = _nanovm_code.charCodeAt(_nanovm_ip);
        ++_nanovm_ip;

        if (opcode >= 0x61 && opcode <= 0x7A) {
            // Small letters.
            _nanovm_stack.push(_nanovm_registers[opcode - 0x61]);
        }
        else if (opcode >= 0x41 && opcode <= 0x5A) {
            // Capital letters.
            var x = _nanovm_stack.pop();
            var regno = opcode - 0x41;
            _nanovm_registers[regno] = x;
            _nanovm_set_register_callback(regno, x);
        }
        else if (opcode >= 0x30 && opcode <= 0x39) {
            // Digits.
            var number = opcode - 0x30;
            for (; _nanovm_ip < _nanovm_code.length; ++_nanovm_ip) {
                var number2 = _nanovm_code.charCodeAt(_nanovm_ip) - 0x30;
                if (number2 >= 0 && number2 <= 9) {
                    number = number * 10 + number2;
                }
                else {
                    break;
                }
            }
            _nanovm_stack.push(number);
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
                    var arg2 = _nanovm_stack.pop();
                    var arg1 = _nanovm_stack.pop();
                    _nanovm_stack.push(arg1 * arg2);
                    break;
                case 0x2B: // '+'
                    var arg2 = _nanovm_stack.pop();
                    var arg1 = _nanovm_stack.pop();
                    _nanovm_stack.push(arg1 + arg2);
                    break;
                case 0x2D: // '-'
                    var arg2 = _nanovm_stack.pop();
                    var arg1 = _nanovm_stack.pop();
                    _nanovm_stack.push(arg1 - arg2);
                    break;
                case 0x2F: // '/'
                    var arg2 = _nanovm_stack.pop();
                    var arg1 = _nanovm_stack.pop();
                    _nanovm_stack.push(arg1 / arg2);
                    break;
                default:
                    _nanovm_println_callback('Unknown opcode: ' + opcode);
            }
        }
        return 1;
    }
    else {
        return 0;
    }
}
