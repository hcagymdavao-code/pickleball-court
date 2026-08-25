(function () {
  "use strict";

  var ADMIN_PASSWORD = "hca12345";
  var CONFIG = window.PICKLEBALL_CONFIG || {};
  var WEB_APP_URL = CONFIG.sheetsWebAppUrl || "";
  var SPREADSHEET_URL = CONFIG.spreadsheetUrl || "";
  var SPREADSHEET_ID = (function extractSpreadsheetId(url) {
    var match = String(url || "").match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : "";
  })(SPREADSHEET_URL);
  var REQUIRED_SCHEMA = 2;
  var COURTS = [
    { id: "1", name: "Court 1", wing: "Left Wing", label: "Court 1 (Left Wing)" },
    { id: "2", name: "Court 2", wing: "Right Wing", label: "Court 2 (Right Wing)" }
  ];

  angular
    .module("pickleballApp", [])
    .config([
      "$sceDelegateProvider",
      function ($sceDelegateProvider) {
        $sceDelegateProvider.trustedResourceUrlList([
          "self",
          "https://script.google.com/**",
          "https://script.googleusercontent.com/**",
          "https://opensheet.elk.sh/**",
          "https://docs.google.com/**"
        ]);
      }
    ])
    .controller("SchedulerController", [
    "$http",
    "$httpParamSerializer",
    "$timeout",
    function ($http, $httpParamSerializer, $timeout) {
      var vm = this;

      vm.slots = buildSlots();
      vm.schedules = [];
      vm.dateValue = null;
      vm.selectedDateKey = "";
      vm.selectedDateLabel = "";
      vm.confirmOpen = false;
      vm.pendingAction = null;
      vm.password = "";
      vm.passwordError = "";
      vm.toast = "";
      vm.groupedSchedules = [];
      vm.saving = false;
      vm.loading = false;
      vm.savedSchedulesEnabled = true;
      vm.savedSchedulesVisible = false;
      vm.sheetReady = !!SPREADSHEET_URL;
      vm.sheetConnectionStatus = "connecting";
      vm.sheetConnectionText = "Connecting";
      vm.accessError = "";
      vm.accessErrorTitle = "";
      vm.spreadsheetUrl = SPREADSHEET_URL;
      vm.courts = COURTS;

      vm.onDateChange = onDateChange;
      vm.toggleCourt = toggleCourt;
      vm.clearCourt = clearCourt;
      vm.onCourtInputClick = onCourtInputClick;
      vm.requestSave = requestSave;
      vm.requestDeleteDate = requestDeleteDate;
      vm.confirmPassword = confirmPassword;
      vm.closeModal = closeModal;
      vm.loadSavedDate = loadSavedDate;

      boot();

      function boot() {
        setSheetConnectionStatus(WEB_APP_URL && navigator.onLine !== false ? "connecting" : "offline");

        if (!WEB_APP_URL) {
          vm.sheetReady = false;
          vm.groupedSchedules = [];
          vm.loading = false;
          return;
        }

        if (!SPREADSHEET_ID) {
          vm.sheetReady = false;
          setSheetConnectionStatus("offline");
          return;
        }

        loadSchedules();
        return checkSheetConnection();
      }

      function checkSheetConnection() {
        if (!WEB_APP_URL || navigator.onLine === false) {
          setSheetConnectionStatus("offline");
          return Promise.resolve(false);
        }

        setSheetConnectionStatus("connecting");
        return sheetsCall("list", { date: "0000-00-00" })
          .then(function () {
            setSheetConnectionStatus("online");
            return true;
          })
          .catch(function () {
            setSheetConnectionStatus("offline");
            return false;
          });
      }

      function loadSchedules() {
        vm.loading = true;
        setSheetConnectionStatus("connecting");
        return sheetsCall("list").then(
          function (res) {
            if (res && res.ok) {
              applyLoaded(res.schedules || [], res);
              setSheetConnectionStatus("online");
              return;
            }
            return loadFromShareLink();
          },
          function () {
            return loadFromShareLink();
          }
        );
      }

      function loadFromShareLink() {
        return $http
          .get("https://opensheet.elk.sh/" + SPREADSHEET_ID + "/Schedules")
          .catch(function () {
            return $http.get(
              "https://opensheet.elk.sh/" + SPREADSHEET_ID + "/Sheet1"
            );
          })
          .then(function (res) {
            applyLoaded(normalizeSheetRows(res.data));
            setSheetConnectionStatus("online");
          })
          .catch(function () {
            applyLoaded([]);
            setSheetConnectionStatus("offline");
          });
      }

      function applyLoaded(items, scriptRes) {
        vm.accessError = "";
        vm.accessErrorTitle = "";
        vm.schedules = sortSchedules(sanitizeSchedules(items));
        refreshGroups();
        if (vm.selectedDateKey) {
          applyDateToSlots();
        }
        vm.loading = false;
        setSheetConnectionStatus("online");
        if (scriptRes && !hasCurrentSchema(scriptRes)) {
          vm.accessErrorTitle = "Apps Script is still the old version";
          vm.accessError = outdatedScriptMessage();
        }
      }

      function hasCurrentSchema(res) {
        return Number(res && res.schemaVersion) >= REQUIRED_SCHEMA;
      }

      function outdatedScriptMessage() {
        return "Pasting Code.gs only updates the editor. The website still talks to the last deployed web app, which writes the old 6-column format (notes land in Court). In Apps Script: Deploy → Manage deployments → pencil on the existing web app → Version = New version → Deploy. Do not create a new deployment, or the URL in js/config.js will no longer match.";
      }

      function sanitizeSchedules(items) {
        if (!angular.isArray(items)) {
          return [];
        }
        return items.filter(function (item) {
          var date = String((item && item.date) || "").slice(0, 10);
          var time = extractStartTime(String((item && item.time) || ""));
          var timeLabel = String((item && item.timeLabel) || "");
          return (
            /^\d{4}-\d{2}-\d{2}$/.test(date) &&
            /^\d{1,2}:\d{2}/.test(time) &&
            date !== "Date" &&
            time !== "Time" &&
            timeLabel !== "Slot"
          );
        });
      }

      function buildSlots() {
        var slots = [];
        var i;
        for (i = 0; i < 24; i += 1) {
          var hour = (6 + i) % 24;
          slots.push({
            hour: hour,
            time: pad(hour) + ":00",
            label: formatHour(hour) + " – " + formatHourEnd(hour),
            courts: {
              "1": { selected: false, notes: "" },
              "2": { selected: false, notes: "" }
            }
          });
        }
        return slots;
      }

      function onDateChange() {
        var key = getInputDateKey();
        if (!key) {
          vm.selectedDateKey = "";
          vm.selectedDateLabel = "";
          resetSlots();
          return;
        }

        vm.selectedDateKey = key;
        vm.selectedDateLabel = formatLongDate(key);
        applyDateToSlots();
      }

      function toggleCourt(slot, courtId) {
        var cell = slot.courts[courtId];
        cell.selected = !cell.selected;
        if (!cell.selected) {
          cell.notes = "";
        }
      }

      function clearCourt(event, slot, courtId) {
        if (event) {
          event.stopPropagation();
        }
        var cell = slot.courts[courtId];
        cell.selected = false;
        cell.notes = "";
      }

      function onCourtInputClick(event, slot, courtId) {
        event.stopPropagation();
        if (!slot.courts[courtId].selected) {
          toggleCourt(slot, courtId);
        }
      }

      function requestSave() {
        if (!vm.selectedDateKey) {
          return;
        }
        vm.pendingAction = { type: "save" };
        openModal();
      }

      function requestDeleteDate(dateKey) {
        if (!vm.savedSchedulesEnabled) {
          return;
        }
        vm.pendingAction = { type: "delete", dateKey: dateKey };
        openModal();
      }

      function confirmPassword() {
        if (vm.password !== ADMIN_PASSWORD) {
          vm.passwordError = "Incorrect password.";
          return;
        }

        var action = vm.pendingAction;
        var password = vm.password;
        closeModal();

        if (action && action.type === "delete") {
          deleteDate(action.dateKey, password);
        } else {
          saveCurrentDate(password);
        }
      }

      function saveCurrentDate(password) {
        var slots = [];
        vm.slots.forEach(function (slot) {
          COURTS.forEach(function (court) {
            var cell = slot.courts[court.id];
            if (cell.selected) {
              slots.push({
                id: createId(),
                time: slot.time,
                timeLabel: slot.label,
                court: court.id,
                courtLabel: court.label,
                notes: (cell.notes || "").trim(),
                savedAt: new Date().toISOString()
              });
            }
          });
        });

        vm.saving = true;
        sheetsCall("save", {
          password: password,
          date: vm.selectedDateKey,
          slots: JSON.stringify(slots)
        }).then(
          function (res) {
            vm.saving = false;
            if (!res || !res.ok) {
              vm.accessError = (res && res.error) || writeAccessMessage();
              showToast(vm.accessError);
              return;
            }
            if (!hasCurrentSchema(res)) {
              vm.accessErrorTitle = "Apps Script is still the old version";
              vm.accessError = outdatedScriptMessage();
              showToast(vm.accessError);
              return;
            }
            vm.accessError = "";
            vm.accessErrorTitle = "";
            showToast("Schedule saved to Google Sheet.");
            loadSchedules();
          },
          function (err) {
            vm.saving = false;
            vm.accessError = writeAccessMessage();
            showToast(err || vm.accessError);
          }
        );
      }

      function deleteDate(dateKey, password) {
        vm.saving = true;
        sheetsCall("delete", {
          password: password,
          date: dateKey
        }).then(
          function (res) {
            vm.saving = false;
            if (!res || !res.ok) {
              showToast((res && res.error) || "Delete failed.");
              return;
            }
            vm.accessError = "";
            showToast("Schedule removed from Google Sheet.");
            loadSchedules();
          },
          function (err) {
            vm.saving = false;
            vm.accessError = writeAccessMessage();
            showToast(err || vm.accessError);
          }
        );
      }

      function writeAccessMessage() {
        return "Google blocked the save because the script still requires a Google login. In Apps Script: Deploy → Manage deployments → edit → Who has access = Anyone. The share link stays the database. Users of this site will not sign in to Google.";
      }

      function normalizeSheetRows(rows) {
        if (!angular.isArray(rows)) {
          return [];
        }
        return rows
          .map(function (row) {
            var rawTime = String(row.Time || row.time || row.Slot || row.slot || "");
            var time = normalizeTimeValue(rawTime);
            var hour = Number(time.split(":")[0]);
            var court = String(row.Court || row.court || "1");
            if (court !== "2") {
              court = "1";
            }
            return {
              date: String(row.Date || row.date || "").slice(0, 10),
              time: time,
              timeLabel: String(row.Slot || row.slot || row.timeLabel || rawTime),
              court: court,
              courtLabel:
                String(row["Court Label"] || row.courtLabel || "") ||
                courtLabel(court),
              notes: String(row.Notes || row.notes || ""),
              savedAt: String(row["Saved At"] || row.savedAt || ""),
              id: String(row.Id || row.id || ""),
              hour: isNaN(hour) ? 0 : hour
            };
          })
          .filter(function (item) {
            return (
              /^\d{4}-\d{2}-\d{2}$/.test(item.date) &&
              /^\d{1,2}:\d{2}/.test(item.time) &&
              item.date !== "Date" &&
              item.time !== "Time" &&
              item.timeLabel !== "Slot"
            );
          });
      }

      function normalizeTimeValue(value) {
        return extractStartTime(String(value || "").trim());
      }

      function extractStartTime(value) {
        var text = String(value || "").trim();
        var match = text.match(/(\d{1,2}):(\d{2})/);
        if (!match) {
          return text;
        }
        var hour = Number(match[1]);
        var minute = match[2];
        return pad(hour) + ":" + minute;
      }

      function sheetsCall(action, extra) {
        if (!WEB_APP_URL) {
          return rejected(writeAccessMessage());
        }

        var params = angular.extend({ action: action }, extra || {});
        return $http
          .jsonp(WEB_APP_URL, {
            jsonpCallbackParam: "callback",
            params: params
          })
          .then(function (res) {
            return res.data;
          })
          .catch(function () {
            var query = $httpParamSerializer(params);
            return $http.get(WEB_APP_URL + "?" + query).then(
              function (res) {
                return res.data;
              },
              function () {
                throw "Could not reach Google Sheets.";
              }
            );
          });
      }

      function rejected(message) {
        setSheetConnectionStatus("offline");
        return {
          then: function (ok, fail) {
            if (fail) {
              fail(message);
            }
          }
        };
      }

      function setSheetConnectionStatus(status) {
        var normalized = status === "online" || status === "offline" ? status : "connecting";
        vm.sheetConnectionStatus = normalized;
        vm.sheetConnectionText = normalized === "online" ? "Online" : normalized === "offline" ? "Offline" : "Connecting";
      }

      function applyDateToSlots() {
        resetSlots();
        vm.schedules.forEach(function (item) {
          if (item.date !== vm.selectedDateKey) {
            return;
          }
          var normalizedItemTime = normalizeTimeValue(item.time);
          vm.slots.forEach(function (slot) {
            if (normalizeTimeValue(slot.time) !== normalizedItemTime) {
              return;
            }
            var court = item.court === "2" ? "2" : "1";
            slot.courts[court].selected = true;
            slot.courts[court].notes = item.notes || "";
          });
        });
      }

      function resetSlots() {
        vm.slots.forEach(function (slot) {
          COURTS.forEach(function (court) {
            slot.courts[court.id].selected = false;
            slot.courts[court.id].notes = "";
          });
        });
      }

      function courtLabel(courtId) {
        var found = COURTS.filter(function (court) {
          return court.id === String(courtId);
        })[0];
        return found ? found.label : "Court 1 (Left Wing)";
      }

      function loadSavedDate(dateKey) {
        if (!vm.savedSchedulesEnabled) {
          return;
        }
        vm.dateValue = parseDateKey(dateKey);
        $timeout(onDateChange);
      }

      function refreshGroups() {
        var groups = {};
        vm.schedules.forEach(function (item) {
          if (!groups[item.date]) {
            groups[item.date] = {
              date: item.date,
              label: formatLongDate(item.date),
              items: []
            };
          }
          groups[item.date].items.push(item);
        });

        vm.groupedSchedules = Object.keys(groups)
          .sort()
          .map(function (key) {
            groups[key].items.sort(function (a, b) {
              var timeDiff =
                slotOrder(Number(itemHour(a))) - slotOrder(Number(itemHour(b)));
              if (timeDiff !== 0) {
                return timeDiff;
              }
              return String(a.court || "1") < String(b.court || "1") ? -1 : 1;
            });
            return groups[key];
          });
      }

      function itemHour(item) {
        if (typeof item.hour === "number") {
          return item.hour;
        }
        return Number(String(item.time || "0").split(":")[0]);
      }

      function openModal() {
        vm.password = "";
        vm.passwordError = "";
        vm.confirmOpen = true;
        $timeout(function () {
          var input = document.getElementById("confirm-password");
          if (input) {
            input.focus();
            input.select();
          }
        }, 50);
      }

      function closeModal() {
        vm.confirmOpen = false;
        vm.pendingAction = null;
        vm.password = "";
        vm.passwordError = "";
      }

      function showToast(message) {
        vm.toast = message;
        $timeout(function () {
          if (vm.toast === message) {
            vm.toast = "";
          }
        }, 3200);
      }

      function createId() {
        return Date.now() + "-" + Math.random().toString(36).slice(2, 8);
      }

      function pad(value) {
        return (value < 10 ? "0" : "") + value;
      }

      function formatHour(hour) {
        var suffix = hour >= 12 ? "PM" : "AM";
        var display = hour % 12;
        if (display === 0) {
          display = 12;
        }
        return display + ":00 " + suffix;
      }

      function formatHourEnd(hour) {
        var endHour = (hour + 1) % 24;
        var endSuffix = endHour >= 12 ? "PM" : "AM";
        var display = endHour % 12;
        if (display === 0) {
          display = 12;
        }
        return display + ":00 " + endSuffix;
      }

      function slotOrder(hour) {
        return (hour + 18) % 24;
      }

      function getInputDateKey() {
        var input = document.getElementById("schedule-date");
        return input && input.value ? input.value : "";
      }

      function parseDateKey(key) {
        var parts = key.split("-");
        return new Date(
          Number(parts[0]),
          Number(parts[1]) - 1,
          Number(parts[2]),
          12,
          0,
          0
        );
      }

      function formatLongDate(dateKey) {
        var date = parseDateKey(dateKey);
        var days = [
          "Sunday",
          "Monday",
          "Tuesday",
          "Wednesday",
          "Thursday",
          "Friday",
          "Saturday"
        ];
        var months = [
          "January",
          "February",
          "March",
          "April",
          "May",
          "June",
          "July",
          "August",
          "September",
          "October",
          "November",
          "December"
        ];
        return (
          days[date.getDay()] +
          ", " +
          months[date.getMonth()] +
          " " +
          date.getDate() +
          ", " +
          date.getFullYear()
        );
      }

      function sortSchedules(items) {
        return items.slice().sort(function (a, b) {
          if (a.date !== b.date) {
            return a.date < b.date ? -1 : 1;
          }
          var timeDiff = slotOrder(itemHour(a)) - slotOrder(itemHour(b));
          if (timeDiff !== 0) {
            return timeDiff;
          }
          return String(a.court || "1") < String(b.court || "1") ? -1 : 1;
        });
      }
    }
  ]);
})();
