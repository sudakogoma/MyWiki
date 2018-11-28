function FindProxyForURL(url, host) {
  if (isInNet(host, "127.0.0.1", "255.255.255.255")) {
        return "PROXY 127.0.0.1:8080";
    }
  else {
    return "DIRECT";
  }
}
