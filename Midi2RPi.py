import sys, time, socket
#from win32midi2 import *
#from midiinlib import *
from ctypes import *
from ctypes.wintypes import *

winmm = windll.LoadLibrary("winmm")
DWORD_PTR = ctypes.POINTER(DWORD)
CALLBACK_FUNCTION = 0x30000  # winmm function type

#Command Line Arguments (3 max)
arg1 = "help"
arg2 = ""
arg3 = ""
midi_in_device = 9
midi_out_device = 9
number_of_arguments = len(sys.argv)
if number_of_arguments > 1:
    arg1 = sys.argv[1]
    if str.isnumeric(arg1):
        midi_in_device = int(arg1)
if number_of_arguments > 2:
    arg2 = sys.argv[2]
    if str.isnumeric(arg2):
        midi_out_device = int(arg2)
if number_of_arguments > 3:
    arg3 = sys.argv[3]

#Flags
#MHDR_DONE = 0x00000001
#MHDR_PREPARED = 0x00000002
#MHDR_INQUEUE = 0x00000004
#MHDR_ISSTRM = 0x00000008

# callback msg_types
INPORT_CLOSED = 0x03C0
INPORT_OPEN = 0x03C1
MESSAGE_IN = 0x03C3
SYSEX_IN = 0x03C4
SYSEX_BUFFER_LENGTH = 256

txqueue = [0] * 1024
txqueue_in = 0
txqueue_out = 0
txqueue_size = 0
txqueue_max = 0

sysex_received = False
sysex_in_progress = False
sysex_index = 0
udp_sysex_count = 0
midi_sysex_count = 0
DEBUG_COLUMNS = 2

udp_buffer_element = bytearray(64)
udp_buffer_queue = [udp_buffer_element] * 1024
udp_buffer_queue_in = 0
udp_buffer_queue_out = 0
udp_buffer_queue_size = 0
udp_buffer_queue_max = 0

# CLASS DEFINITIONS

class MIDIINCAPSA(Structure):
    _fields_ = [("wMid", WORD),
                ("wPid", WORD),
                ("vDriverVersion", UINT),
                ("_szPname", c_byte * 32),
                ("dwSupport", DWORD),
                ]
    def __getattr__(self, name):
        if name == 'szPname':
            for i in range(32):
                if self._szPname[i] == 0:
                    EOStr = i
                    break
            return ''.join(map(chr, self._szPname[:EOStr]))
        else:
            return Structure.name

class MIDIOUTCAPSA(Structure):
    _fields_ = [("wMid", WORD),
                ("wPid", WORD),
                ("vDriverVersion", UINT),
                # ("szPname", CHAR*32),
                ("_szPname", c_byte * 32),
                ("wTechnology", WORD),
                ("wVoices", WORD),
                ("wNotes", WORD),
                ("wChannelMask", WORD),
                ("dwSupport", DWORD),
                ]
    def __getattr__(self, name):
        if name == 'szPname':
            for i in range(32):
                if self._szPname[i] == 0:
                    EOStr = i
                    break
            return ''.join(map(chr, self._szPname[:EOStr]))
        else:
            return Structure.name

class midinote():
    "A midi note"
    def __init__(self, msg_type, raw,time,instanceData):
        self.raw = raw
        #self.command = raw & 0xFF
        #self.data1 = (raw >> 8) & 0xFF
        #self.data2 = (raw >> 16) & 0xFF
        self.self = self
        self.msg_type = msg_type
        self.time = time
        self.instanceData = instanceData        

class midiIn ():
    "A inbound midi connection"
    def start(self, midiDevID=-1):
        if (self.status == 1):
            if (midiDevID != -1 and midiDevID != self.midiDevID):
                self.stop()
            else:
                return
        if (midiDevID == -1):
            midiDevID = self.midiDevID
        else:
            self.midiDevID = midiDevID
        self.error = winmm.midiInOpen(byref(self.midiInID), # reference
                 c_int(midiDevID),  # Midi device to use
                 self.midi_get,     # callback function
                 c_int(0),          # instance data
                 CALLBACK_FUNCTION) # Callback flag
        winmm.midiInStart(self.midiInID)
        self.status = 1
        self.device_id = self.midiInID
    def __init__(self, midiDevID=0 ,fun=0):
        self.midiInID = c_long()
        self.device_id = c_long()
        self.midiDevID = midiDevID
        self.status = 0 # 0=stopped, 1=started
        if (fun == 0):
            def defaultProcFun(note):
                pass
            self.fun = defaultProcFun
        else:
            self.fun=fun
        self.CMPFUNC = WINFUNCTYPE(None, c_long, UINT, DWORD, DWORD, DWORD)
        def MidiSigRec(midiInID,msg_type,instanceData,raw,time):
            note = midinote(msg_type, raw, time, instanceData)
            self.fun(note)
        self.MidiSigRec = MidiSigRec
        self.midi_get = self.CMPFUNC(self.MidiSigRec)
        self.start(midiDevID)
    def suspend (self):
        winmm.midiInStop(self.midiInID)
    def reset (self):
        winmm.midiInReset(self.midiInID)
    def restart (self):
        winmm.midiInStart(self.midiInID)
    def stop (self):
        winmm.midiInStop(self.midiInID)
        winmm.midiInClose(self.midiInID)
        self.midiInID = c_long(0)
        self.status = 0
        
class MIDIHDR(Structure):
    pass
MIDIHDR._fields_ = [("lpData", LPSTR),
                    ("dwBufferLength", DWORD),
                    ("dwBytesRecorded", DWORD),
                    ("dwUser", DWORD_PTR),
                    ("dwFlags", DWORD),
                    ("lpNext", POINTER(MIDIHDR)),
                    ("reserved", DWORD),
                    ("dwOffset", DWORD),
                    ("dwReserved", DWORD_PTR * 4)
                    ]


# FUNCTION DEFINITIONS

def midiInGetDevCapsA(uDeviceID, lpMidiOutCaps):
    return winmm.midiInGetDevCapsA(uDeviceID, byref(lpMidiOutCaps),
                                 ctypes.sizeof(lpMidiOutCaps))

def midiOutGetDevCapsA(uDeviceID, lpMidiOutCaps):
    return winmm.midiOutGetDevCapsA(uDeviceID, byref(lpMidiOutCaps),
                                  ctypes.sizeof(lpMidiOutCaps))

def midiInPrepareHeader(hmi, lpMidiInHdr):
    return winmm.midiInPrepareHeader(hmi, byref(lpMidiInHdr),
                                    ctypes.sizeof(lpMidiInHdr))


def midiInAddBuffer(hmi, lpMidiInHdr):
    return winmm.midiInAddBuffer(hmi, byref(lpMidiInHdr),
                                    ctypes.sizeof(lpMidiInHdr))


def midiInUnprepareHeader(hmi, lpMidiInHdr):
    return winmm.midiInUnprepareHeader(hmi, byref(lpMidiInHdr),
                                      ctypes.sizeof(lpMidiInHdr))


def midiOutPrepareHeader(hmo, lpMidiOutHdr):
    return winmm.midiOutPrepareHeader(hmo, byref(lpMidiOutHdr),
                                    ctypes.sizeof(lpMidiOutHdr))


def midiOutLongMsg(hmo, lpMidiOutHdr):
    return winmm.midiOutLongMsg(hmo, byref(lpMidiOutHdr),
                              ctypes.sizeof(lpMidiOutHdr))


def midiOutUnprepareHeader(hmo, lpMidiOutHdr):
    return winmm.midiOutUnprepareHeader(hmo, byref(lpMidiOutHdr),
                                      ctypes.sizeof(lpMidiOutHdr))

def midiOutOpen(dev=0, a=0, b=0, c=0):
    h = c_int()
    r = winmm.midiOutOpen(byref(h), dev, 0, 0, 0)
    return r, h

def midiOutClose(h):
    return winmm.midiOutClose(h)

def midiOutShortMsg(h, dwMsg):
    return winmm.midiOutShortMsg(h, dwMsg)

def MidiCallback(midi_in):
    global txqueue, txqueue_in, txqueue_size, hdr1, sysex_received, midi_sysex_count
    if midi_in.msg_type == MESSAGE_IN:
        if midi_in.raw & 0xFF != 0xF8 and midi_in.raw & 0xFF != 0xFE:
            txqueue[txqueue_in] = midi_in.raw
            txqueue_in = txqueue_in + 1
            if txqueue_in >= len(txqueue):
                txqueue_in = 0
            txqueue_size = txqueue_size + 1
    elif midi_in.msg_type == SYSEX_IN:
        byte_count = 0
        midi_sysex_count = midi_sysex_count + 1
        for sysex_count in range(0, hdri.dwBytesRecorded):
            byte_count = byte_count + 1
            if byte_count == 3:
                txqueue[txqueue_in] = txqueue[txqueue_in] | ((bi[sysex_count] & 0xFF ) << 16)
                txqueue_in = txqueue_in + 1
                if txqueue_in >= len(txqueue):
                    txqueue_in = 0
                txqueue_size = txqueue_size + 1
                byte_count = 0
            elif byte_count == 2:
                txqueue[txqueue_in] = txqueue[txqueue_in] | ((bi[sysex_count] & 0xFF ) << 8)
            elif byte_count == 1:
                txqueue[txqueue_in] = bi[sysex_count] & 0xFF 
        if byte_count != 0:
            txqueue_in = txqueue_in + 1
            if txqueue_in >= len(txqueue):
                txqueue_in = 0
            txqueue_size = txqueue_size + 1
        sysex_received = True
    elif midi_in.msg_type == INPORT_OPEN:
        print("CALLBACK MESSAGE ... Port Opened")
    elif midi_in.msg_type == INPORT_CLOSED:
        print("CALLBACK MESSAGE ... Port Closed")
    else:
        print("CALLBACK ERROR ! ... msg_type =", midi_in.msg_type)

command_list = [0x80, 0x90, 0xA0, 0xB0, 0xC0, 0xD0, 0xE0,
                0xF0, 0xF1, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7,
                0xF8, 0xF9, 0xFA, 0xFB, 0xFC, 0xFD, 0xFE, 0xFF]

cable_list   = [0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E,
                0x04, 0x02, 0x03, 0x02, 0x0F, 0x0F, 0x05, 0x05,
                0x05, 0x0F, 0x05, 0x05, 0x05, 0x0F, 0x05, 0x05]

def cable_lookup(command_byte, data_byte1, data_byte2):
    cable_byte = 0x04
    if command_byte >= 0x80:
        if command_byte < 0xF0:
            command_byte = command_byte & 0xF0
        for i in range(0, len(command_list)):
            if command_byte == command_list[i]:
                cable_byte = cable_list[i]
    if data_byte1 == 0xF7:
        cable_byte = 0x06
    if data_byte2 == 0xF7:
        cable_byte = 0x07
    return cable_byte

hex_list = ["0x00", "0x01", "0x02", "0x03", "0x04", "0x05", "0x06", "0x07",
            "0x08", "0x09", "0x0a", "0x0b", "0x0c", "0x0d", "0x0e", "0x0f"]

def print_hex(hex_val):
    if hex_val < 0x10:
        print(hex_list[hex_val], end=' ')
    else:
        print(hex(hex_val), end=' ')

def print_usb_array(usb_array, array_length):
    column = 1
    first_line = True
    i = 1
    while i < array_length:
        if usb_array[i] != 0x00:
            if not first_line and column == 1:
                print("    ", end='')
            print_hex(usb_array[i - 1])
            print_hex(usb_array[i])
            print_hex(usb_array[i + 1])
            print_hex(usb_array[i + 2])
            if column >= DEBUG_COLUMNS:
                print()
                column = 1
                first_line = False
            else:
                print("    ", end='')
                column = column + 1
        i = i + 4
    if column > 1:
        print()

def midi_to_udp(size_to_send):   # formats data into 64 byte usb midi spec
    global txqueue, txqueue_out, txqueue_size
    udp_out_data = bytearray(64)
    udp_index = 0
    for i in range(0, size_to_send):
        command_byte = txqueue[txqueue_out] & 0xFF
        data_byte1 = (txqueue[txqueue_out] >> 8) & 0xFF
        data_byte2 = (txqueue[txqueue_out] >> 16) & 0xFF
        udp_out_data[udp_index] = cable_lookup(command_byte, data_byte1, data_byte2)
        udp_index = udp_index + 1
        udp_out_data[udp_index] = command_byte
        udp_index = udp_index + 1
        udp_out_data[udp_index] = data_byte1
        udp_index = udp_index + 1
        udp_out_data[udp_index] = data_byte2
        udp_index = udp_index + 1
        txqueue_out = txqueue_out + 1
        if txqueue_out >= len(txqueue):
            txqueue_out = 0
        txqueue_size = txqueue_size - 1
    try:
        udp_out.sendto(udp_out_data, output_port)
        #print_usb_array(udp_out_data, len(udp_out_data))
    except OSError as neterror:
        print("Error Sending UDP", neterror)

def send_sysex(buffer_length):
    global bo, hdro, udp_sysex_count
    hdro.lpData = bytes(bo)
    hdro.dwBufferLength = buffer_length
    midiOutLongMsg(h, hdro)
    udp_sysex_count = udp_sysex_count + 1
    #print_usb_array(bo, buffer_length + 1)
    
def net_to_midi(usb_data):
    global bo, udp_sysex_count, midi_sysex_count, sysex_in_progress, sysex_index
    global udp_buffer_queue, udp_buffer_queue_in, udp_buffer_queue_out
    global udp_buffer_queue_size, udp_buffer_queue_max
    if arg1 == "debug" or arg2 == "debug" or arg3 == "debug":
        udp_buffer_queue[udp_buffer_queue_in] = usb_data
        udp_buffer_queue_in = udp_buffer_queue_in + 1
        if udp_buffer_queue_in >= len(udp_buffer_queue):
            udp_buffer_queue_in = 0
        udp_buffer_queue_size = udp_buffer_queue_size + 1
    i = 1
    while i < len(usb_data):
        if usb_data[i] == 0xF0 or sysex_in_progress:
            sysex_in_progress = True
            bo[sysex_index] = usb_data[i]
            sysex_index = sysex_index + 1
            if usb_data[i] == 0xF7:
                sysex_in_progress = False
                send_sysex(sysex_index - 1)
                sysex_index = 0
            else:
                bo[sysex_index] = usb_data[i + 1]
                sysex_index = sysex_index + 1
                if usb_data[i + 1] == 0xF7:
                    sysex_in_progress = False
                    send_sysex(sysex_index - 1)
                    sysex_index = 0
                else:
                    bo[sysex_index] = usb_data[i + 2]
                    sysex_index = sysex_index + 1
                    if usb_data[i + 2] == 0xF7:
                        sysex_in_progress = False
                        send_sysex(sysex_index - 1)
                        sysex_index = 0
        elif usb_data[i] & 0x80 != 0x00:
            if sysex_in_progress:
                sysex_in_progress = False
                sysex_index = 0
                print("Sysex Terminated Unexpectedly !")
            if usb_data[i] != 0xF8 and usb_data[i] != 0xFE:
                midiOutShortMsg(h, usb_data[i] | (usb_data[i + 1] << 8) | (usb_data[i + 2] << 16))
                if usb_data[i - 1] == 0x0B and usb_data[i + 1] == 0x40 and usb_data[i + 2] == 0x7F:
                    print("UDP Sysex =", udp_sysex_count, " ... MIDI Sysex =", midi_sysex_count)
                    udp_sysex_count = 0
                    midi_sysex_count = 0
                    if arg1 == "debug" or arg2 == "debug" or arg3 == "debug":
                        print("UDP Input Buffer ...")
                        while udp_buffer_queue_size > 0:
                            print_usb_array(udp_buffer_queue[udp_buffer_queue_out], len(udp_buffer_queue[udp_buffer_queue_out]))
                            udp_buffer_queue_out = udp_buffer_queue_out + 1
                            if udp_buffer_queue_out >= len(udp_buffer_queue):
                                udp_buffer_queue_out = 0
                            udp_buffer_queue_size = udp_buffer_queue_size - 1
        i = i + 4


# MAIN PROGRAM

if arg1 == "help" or arg2 == "help"or arg3 == "help":
    print("Command Line Arguments (3 max) ... n m list select debug help")
    print()
    print("n      ... Midi Input  Port (default = 9 Maple Midi In: Port 2)")
    print("m      ... Midi Output Port (default = 9 Maple Midi Out: Port 1)")
    print("list   ... Lists the available Midi Ports")
    print("select ... Allows interactive selection of Ports. Overrides n and m values")
    print("debug  ... Enables debug output (displayed after Sustain press)")
    print("           n and m must occur in positions 1 & 2")

if arg1 == "debug" or arg2 == "debug" or arg3 == "debug":
    print()
    print("Debug Enabled ... Press Sustain Pedal to display Debug Buffer")

incap = MIDIINCAPSA()
if arg1 == "list" or arg2 == "list" or arg3 == "list":
    print()
    print("Input Devices ...")
    for i in range(0, winmm.midiInGetNumDevs()):
        print(i, end=" ")
        midiInGetDevCapsA(i, incap)
        print(incap.szPname)

if arg1 == "select" or arg2 == "select" or arg3 == "select":
    print()
    while True:
        try:
            midi_in_device = int(input("Enter Input Device  "))
        except:
            print("INVALID ENTRY !")
        else:
            if midi_in_device < 0 or midi_in_device >= winmm.midiInGetNumDevs():
                print("INVALID DEVICE !")
            else:
                break

cap = MIDIOUTCAPSA()
if arg1 == "list" or arg2 == "list" or arg3 == "list":
    print()
    print("Output Devices ...")
    for i in range(0, winmm.midiOutGetNumDevs()):
        print(i, end=" ")
        midiOutGetDevCapsA(i, cap)
        print(cap.szPname)

if arg1 == "select" or arg2 == "select" or arg3 == "select":
    print()
    while True:
        try:
            midi_out_device = int(input("Enter Output Device  "))
        except:
            print("INVALID ENTRY !")
        else:
            if midi_out_device < 0 or midi_out_device >= winmm.midiOutGetNumDevs():
                print("INVALID DEVICE !")
            else:
                break

print()
print("Opening Input Device", midi_in_device, end=" : ")
midiInGetDevCapsA(midi_in_device, incap)
print(incap.szPname)

inport = midiIn(midi_in_device, MidiCallback)

hdri = MIDIHDR()
bi = bytes(SYSEX_BUFFER_LENGTH)
hdri.lpData = bi
hdri.dwBufferLength = SYSEX_BUFFER_LENGTH
hdri.dwFlags = 0
midiInPrepareHeader(inport.device_id, hdri)
midiInAddBuffer(inport.device_id, hdri)
time.sleep(0.5)

print()
print("Opening Output Device", midi_out_device, end=" : ")
midiOutGetDevCapsA(midi_out_device, cap)
print(cap.szPname)
r, h = midiOutOpen(midi_out_device, 0, 0, 0)
winmm.midiOutReset(h)

hdro = MIDIHDR()
bo = bytearray(SYSEX_BUFFER_LENGTH)
hdro.lpData = bytes(bo)
hdro.dwBufferLength = 0
midiOutPrepareHeader(h, hdro)

print()
print("Setting Up Ethernet")
output_port = ("192.168.0.40", 6666)
try:
    udp_out = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
except OSError as neterror:
    print(neterror)
    print('Failed to create socket')
    sys.exit()

rpi_eth = ('', 5555)
try:
    netmidi = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
except OSError as neterror:
    print(neterror)
    print('Failed to create socket')
    sys.exit()
    
netmidi.setblocking(False)

try:
    netmidi.bind(rpi_eth)
    # fails here when more than one instance is run
    # only allows a single binding to a port
except OSError as neterror:
    print(neterror)
    print("ETHERNET UNAVAILABLE")
    sys.exit()
else:
    print("ETHERNET DETECTED")

time.sleep(1)
#print("")
#print("DISABLE ONLINE ARMOR FIREWALL TO AVOID DELAYS !")
print()
print("Press CTRL-C to Exit")
print()

while True:
    if sysex_received:
        sysex_received = False
        midiInAddBuffer(inport.device_id, hdri)
        
    if txqueue_size > txqueue_max:
        txqueue_max = txqueue_size
    if txqueue_size > 0:
        if txqueue_size <= 16:
            midi_to_udp(txqueue_size)
        else:
            midi_to_udp(16)
            
    try:
        time.sleep(0.01)
        # allows OS a look-in, also mostly eliminates RPi dropping packets
    except KeyboardInterrupt:
        print("... EXITING")
        break
    
    try:
        try:
            udpdata = netmidi.recvfrom(1024)
        except OSError as neterror:
            if neterror.args[0] != 10035:
                print(neterror)
                break
        else:
             netdata = udpdata[0]
             net_to_midi(netdata)
    except KeyboardInterrupt:
        print("... EXITING")
        break

time.sleep(1)
midiInUnprepareHeader(inport.device_id, hdri)
inport.stop()
midiOutUnprepareHeader(h, hdro)
midiOutClose(h)
print("Midi Ports Closed")
time.sleep(1)
#print()
#print("RE-ENABLE ONLINE ARMOR FIREWALL !")
sys.exit()
