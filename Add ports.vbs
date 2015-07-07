Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile
Dim portTag 
portTag = "Port"
'constant for UDP 17 
Dim UDP
UDP = 17
'constant for TCP 6
Dim TCP
TCP = 6
'Enable ICMP
'Set objICMPSettings = objPolicy.ICMPSettings
'objICMPSettings.AllowInboundEchoRequest = TRUE



Function addPorts(initial, final, portTag, Protocol)
if IsNull(final)then

Set objPort = CreateObject("HNetCfg.FwOpenPort")
		
		objPort.Port = initial
		objPort.Name =  portTag & " " & initial
		objPort.Protocol = Protocol
		objPort.Enabled = FALSE
		Set colPorts = objPolicy.GloballyOpenPorts
		errReturn = colPorts.Add(objPort)
Else

For Port= initial To final Step 1
		Set objPort = CreateObject("HNetCfg.FwOpenPort")
		objPort.Port = Port
		objPort.Name = portTag & " " & Port
		objPort.Protocol = Protocol
		objPort.Enabled = TRUE
		Set colPorts = objPolicy.GloballyOpenPorts
		errReturn = colPorts.Add(objPort)
	Next
End If
End Function
' add parameters for open ports or ranges in Windows xp  (initial Port, final Port, name of the open port "", Choose UDP or TCP  )



Call addPorts(22, null,"SSH_PORT_", TCP)
Call addPorts(23, null,"Telnet_PORT_", TCP)
Call addPorts(135, null,"RPC_PORT_", TCP)
Call addPorts(139, null,"RPC_PORT_", TCP)
Call addPorts(445, null,"RPC_PORT_", TCP)
Call addPorts(1433,1437,"SQL_DBM_PORT_", TCP)
Call addPorts(1521, 1523, "ORACLE_DB_", TCP)
Call addPorts(1526, null, "SQL_DBM_PORT_", TCP)
Call addPorts(1529, null, "SQL_DBM_PORT_", TCP)
Call addPorts(1538, null, "SQL_DBM_PORT_", TCP)
Call addPorts(1548, null, "SQL_DBM_PORT_", TCP)
Call addPorts(49154, null, "WMI_PORT_", TCP)
Call addPorts(49183, null, "WMI_PORT_", TCP)
Call addPorts(49175, null, "WMI_PORT_", TCP)
Call addPorts(5355, null,"RPC_PORT_", TCP)
Call addPorts(50000,60000, "DB2_PORT_", TCP)
Call addPorts(5355, null, "Port_", UDP)
Call addPorts(137,138, "NETBII_", UDP)
Call addPorts(161,162, "SMMP_PORT_", UDP)