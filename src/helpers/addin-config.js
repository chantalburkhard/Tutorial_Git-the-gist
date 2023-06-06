function getConfig() {
  const config = {};

  config.gitHubUserName = Office.context.roamingSettings.get("gitHubUserName");
  config.defaulfGistId = Office.context.roamingSettings.get("defaultGistId");

  return config;
}

function setConfig(config, callback) {
  Office.context.roamingSettings.set("gitHubUserName", config.gitHubUserName);
  Office.context.roamingSettings.set("defaultGistId", config.defaulfGistId);

  Office.context.roamingSettings.saveAsync(callback);
}