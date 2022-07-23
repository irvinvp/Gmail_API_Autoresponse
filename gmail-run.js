const fs = require("fs");
const readline = require("readline");
const { google } = require("googleapis");
const { io } = require("socket.io-client");
var moment = require("moment-timezone");
const socket = io("", {
  transports: ["websocket"],
  query: { type: "query" },
});
let socket_interval;
let geos = {};
let routes = {};
let units = {};
let drivers = {};
socket.on("hello", (...args) => {
  console.log("hello");
  socket.emit(
   "hello",
    "",
    "",
    0,
    ""
  );
});
socket.on("user_info", (...args) => {
  console.log("user_info");
  try {
    clearInterval(socket_interval);
  } catch (e) {}
  socket_interval = setInterval(() => {
    socket.emit("geo_class_list", "");
    socket.emit("route_list", "");
    socket.emit("nick_list", "");
    socket.emit("driver_list", "");
  }, 1 * 60 * 1000);
  socket.emit("geo_class_list", "");
  socket.emit("route_list", "");
  socket.emit("nick_list", "");
  socket.emit("driver_list", "");
});
socket.on("geo_class_list", (...args) => {
  let geo_list = JSON.parse(args[0]);
  let geo_list_ = {};
  for (let x in geo_list) {
    if (geo_list[x][1]) {
      geo_list_[
        geo_list[x][1]
          .toLowerCase()
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
      ] = geo_list[x][0];
    }
  }
  console.log("geo_class_list");
  geos = geo_list_;
});
socket.on("route_list", (...args) => {
  let route_list = args[0];
  let route_list_ = {};
  for (let x in route_list) {
    route_list_[
      route_list[x][1]
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
    ] = route_list[x][0];
  }
  console.log("route_list");
  routes = route_list_;
});
socket.on("nick_list", (...args) => {
  let nick_list = JSON.parse(args[0]);
  let nick_list_ = {};
  for (let x in nick_list) {
    nick_list_[
      nick_list[x]
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
    ] = x;
  }
  console.log("nick_list");
  units = nick_list_;
});
socket.on("driver_list", (...args) => {
  let driver_list = JSON.parse(args[0]);
  let driver_list_ = {};
  for (let x in driver_list) {
    driver_list_[
      driver_list[x][1].name
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
    ] = driver_list[x][0];
  }
  console.log("driver_list");
  drivers = driver_list_;
});

const SCOPES = ["https://www.googleapis.com/auth/gmail.modify"];
const TOKEN_PATH = "token.json";
let gmail;
function run() {
  fs.readFile("credentials.json", (err, content) => {
    if (err) return console.log("Error loading client secret file:", err);
    authorize(JSON.parse(content), listLabels);
  });

  function authorize(credentials, callback) {
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(
      client_id,
      client_secret,
      redirect_uris[0]
    );
    fs.readFile(TOKEN_PATH, (err, token) => {
      if (err) return getNewToken(oAuth2Client, callback);
      oAuth2Client.setCredentials(JSON.parse(token));
      callback(oAuth2Client);
    });
  }

  function getNewToken(oAuth2Client, callback) {
    const authUrl = oAuth2Client.generateAuthUrl({
      access_type: "offline",
      scope: SCOPES,
    });
    console.log("Authorize this app by visiting this url:", authUrl);
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    rl.question("Enter the code from that page here: ", (code) => {
      rl.close();
      oAuth2Client.getToken(code, (err, token) => {
        if (err) return console.error("Error retrieving access token", err);
        oAuth2Client.setCredentials(token);
        fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
          if (err) return console.error(err);
          console.log("Token stored to", TOKEN_PATH);
        });
        callback(oAuth2Client);
      });
    });
  }

  async function listLabels(auth) {
    gmail = google.gmail({ version: "v1", auth });
    let res = await gmail.users.messages.list({
      userId: "me",
      maxResults: 10,
      labelIds: ["UNREAD"],
      q: "",
    });
    for (let x in res.data.messages) {
      let res2 = await gmail.users.messages.get({
        userId: "me",
        id: res.data.messages[x].id,
      });
      for (let x in res2.data.payload.parts) {
        if (res2.data.payload.parts[x].mimeType == "text/plain") {
          let s1 = new Buffer.from(
            res2.data.payload.parts[x].body.data,
            "base64"
          );
          let from = "";
          for (let z in res2.data.payload.headers) {
            if (res2.data.payload.headers[z].name == "From") {
              from = res2.data.payload.headers[z].value;
            }
          }
          let ans = main(s1.toString("utf8"));
          send_mail(from, JSON.stringify(ans));
          gmail.users.messages.modify({
            userId: "me",
            id: res.data.messages[x].id,
            resource: {
              addLabelIds: [],
              removeLabelIds: ["UNREAD"],
            },
          });
          break;
        }
      }
    }
  }
}
setInterval(run, 60 * 1000);
setTimeout(run, 10 * 1000);
function main(data) {
  let data2 = data.toLowerCase().split("\n");
  let end = {};
  for (let x in data2) {
    if (data2[x].includes("unidad")) {
      end.unidad = data2[x]
        .split(":")[1]
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
    }
    if (data2[x].includes("fecha")) {
      end.fecha = data2[x]
        .split(":")[1]
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
    }
    if (data2[x].includes("hora") && data2[x].split(":").length == 3) {
      end.hora =
        data2[x].split(":")[1].trim() + ":" + data2[x].split(":")[2].trim();
    }
    if (data2[x].includes("origen")) {
      end.origen = data2[x]
        .split(":")[1]
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
    }
    if (data2[x].includes("destino")) {
      end.destino = data2[x]
        .split(":")[1]
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
    }
    if (data2[x].includes("operador")) {
      end.operador = data2[x]
        .split(":")[1]
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
    }
    if (data2[x].includes("ruta")) {
      end.ruta = data2[x]
        .split(":")[1]
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
    }
  }
  console.log(end);
  if (
    end.unidad &&
    end.fecha &&
    end.hora &&
    end.origen &&
    end.destino &&
    end.operador &&
    end.ruta
  ) {
    let f;
    if (typeof units[end.unidad] == "undefined") {
      console.log("No existe la unidad");
      return { error: "No existe la unidad" };
    }
    if (typeof geos[end.origen] == "undefined") {
      console.log("No existe el origen");
      return { error: "No existe el origen" };
    }
    if (typeof geos[end.destino] == "undefined") {
      console.log("No existe el destino");
      return { error: "No existe el destino" };
    }
    if (typeof drivers[end.operador] == "undefined") {
      console.log("No existe el operador");
      return { error: "No existe el operador" };
    }
    if (typeof routes[end.ruta] == "undefined" && end.ruta != "") {
      console.log("No existe la ruta");
      return { error: "No existe la ruta" };
    }
    try {
      f = moment(end.fecha + " " + end.hora, "DD/MM/YY HH:mm")
        .tz("America/Mexico_City")
        .unix();
    } catch (e) {
      return { error: "Formato Fecha" };
    }
    socket.emit(
      "travel_save",
      units[end.unidad],
      geos[end.origen],
      geos[end.destino],
      drivers[end.operador],
      routes[end.ruta],
      f,
      [],
      [],
      []
    );
    return { status: "Viaje creado correctamente", datos: end };
  }
  return { error: "Falta datos", data: end };
}

async function send_mail(email, body) {
  var raw = makeBody(
    email,
    "",
    "Automatizacion de viajes " + moment().format("DD/MM/YY HH:mm"),
    body
  );
  let res = await gmail.users.messages.send({
    userId: "me",
    resource: {
      raw: raw,
    },
  });
  console.log(res.data);
}
function makeBody(to, from, subject, message) {
  let str = [
    'Content-Type: text/plain; charset="UTF-8"\n',
    "MIME-Version: 1.0\n",
    "Content-Transfer-Encoding: 7bit\n",
    "to: ",
    to,
    "\n",
    "from: ",
    from,
    "\n",
    "subject: ",
    subject,
    "\n\n",
    message,
  ].join("");
  let encodedMail = new Buffer.from(str)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");
  return encodedMail;
}
