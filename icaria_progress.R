library(redcapAPI)

ReadData <- function(api.url, api.token) {
  rcon <- redcapConnection(api.url, api.token)
  data <- exportRecords(rcon)
}