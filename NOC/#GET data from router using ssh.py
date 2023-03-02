from netmiko import ConnectHandler

router = {
    "device_type": "cisco_ios",
    "ip": "10.109.11.1",
    "username": "NOC.user",
    "password": "EnpoNU@$2023!",
}

net_connect = ConnectHandler(**router)
output = net_connect.send_command("show version")
print(output)
net_connect.disconnect()