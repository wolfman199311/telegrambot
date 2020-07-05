// Copyright 2017, Google, Inc.
// Licensed under the Apache License, Version 2.0 (the 'License');
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//    http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an 'AS IS' BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

// Dialogflow fulfillment getting started guide:
// https://dialogflow.com/docs/how-tos/getting-started-fulfillment

'use strict';

const functions = require('firebase-functions');
const { WebhookClient } = require('dialogflow-fulfillment');
const { Card, Suggestion } = require('dialogflow-fulfillment');
const { Payload } = require("dialogflow-fulfillment");
const apikey = '<apikey>';
const requests = require('request');
var { countries } = require('country-data');
var rp = require('request-promise');
const { error } = require('actions-on-google/dist/common');
var XLSX = require('xlsx');
const OnetWebService = require('./OnetWebService')
const { google } = require('googleapis');
const calendarId = "<calendar ID>"
var appointment_type = 'chatbot';
const serviceAccount = {
  //<service account from google cloud>
}; // Starts with {"type": "service_account",...

process.env.DEBUG = 'dialogflow:debug'; // enables lib debugging statements
const serviceAccountAuth = new google.auth.JWT({
  email: serviceAccount.client_email,
  key: serviceAccount.private_key,
  scopes: 'https://www.googleapis.com/auth/calendar'
});
const calendar = google.calendar('v3');

const timeZone = 'Europe/London';
const timeZoneOffset = '+01:00';

exports.dialogflowFirebaseFulfillment = functions.https.onRequest(async (request, response) => {
  // return new Promise ((resolve, reject) => {});

  const agent = new WebhookClient({ request, response });
  console.log('Dialogflow Request headers: ' + JSON.stringify(request.headers));
  console.log('Dialogflow Request body: ' + JSON.stringify(request.body));

  async function booktrip_activity(agent) {
    try {
      console.log("________book trip activity________")

      //getting parameters from request
      let origin_address = request.body.queryResult.outputContexts[2].parameters["geo-city"];
      var activity = request.body.queryResult.parameters.travel_activities;
      var target_country = request.body.queryResult.outputContexts[1].parameters.target_country;
      let Co2_emission = `${request.body.queryResult.outputContexts[1].parameters.Co2_emission} kg`;
      var diff = request.body.queryResult.outputContexts[1].parameters.diff;
      var homeperc = request.body.queryResult.outputContexts[1].parameters.homeperc;
      var price = request.body.queryResult.outputContexts[1].parameters.price;
      var job_roles = request.body.queryResult.outputContexts[1].parameters.job_roles;
      // let flight_time = request.body.queryResult.outputContexts[1].parameters.flight_time;
      // let time = flight_time.slice(0, 5);
      let time = '4 h 30min';
      // let Co2_emission = '125kg';

      var pwd = 'This trip would require a social security certificate, a PWD registration, and a health screening which I can arrange for you. No work permit and no payroll withholding will be needed.';
      var pwd_p = await need_pwd(activity, target_country, diff);
      var npwd = pwd_p.n_pwd;
      var period = pwd_p.period;
      // determine pwd and remotely(workhome) is needed or not
      switch (homeperc) {
        case 'yes':
          if (npwd) {
            agent.add(pwd);
            // agent.add(`Immi = ${pwd_p.immi}`);
            (pwd_p.immi == 'NP') ? agent.add(`An immigration work permit is not required and pwd is needed if the traveller is there more than ${period} days`) : agent.add(`An Immigration work permit will be required and pwd is needed if the traveller is there more than ${period} days`);
            agent.add(` ${origin_address} flight cost is ${price} USD.`);
            agent.add(new Suggestion('Book trip'));
            agent.add(new Suggestion('Request formal assignment'));
          } else {
            (pwd_p.immi == 'NP') ? agent.add(`An immigration work permit is not required and pwd is needed if the traveller is there more than ${period} days`) : agent.add(`An Immigration work permit will be required  and pwd is needed if the traveller is there more than ${period} days`);
            agent.add(`the cheapest ${origin_address} flight cost is ${price} $.`);
            agent.add(new Suggestion('Book trip'));
            agent.add(new Suggestion('Request formal assignment'));
          }
          break;
        case 'no':
          if (npwd) {
            agent.add(pwd);
            (pwd_p.immi == 'NP') ? agent.add(`An immigration work permit is not required and pwd is needed if the traveller is there more than ${period} days`) : agent.add(`An Immigration work permit will be required and pwd is needed if the traveller is there more than ${period} days`);
            const card = new Card("Alternative options");
            card.setText(`Projects relating to ${job_roles} can usually be completed remotely (100% of historic cases). It could save  ${price} $ in costs, and ${Co2_emission} of CO2 emissions if you avoid this journey.Should I book the trip, explore remote working options, or request a formal assignment to ${origin_address}?`);
            card.setButton({ text: 'Explore Remote Working', url: "https://pwcuk.cobrainer.com/create-job-posting" });
            agent.add(card);
            agent.add(new Suggestion('Book trip'));
            agent.add(new Suggestion('Request formal assignment'));
          } else {
            (pwd_p.immi == 'NP') ? agent.add(`An immigration work permit is not required and pwd is needed if the traveller is there more than ${period} days`) : agent.add(`An Immigration work permit will be required and pwd is needed if the traveller is there more than ${period} days`);
            const card = new Card("Alternative options");
            card.setText(`Projects relating to ${job_roles} can usually be completed remotely(100 % of historic cases).It could save ${price} $ in costs, and ${Co2_emission} of CO2 emissions if you avoid this journey.Should I book the trip, explore remote working options, or request a formal assignment to ${origin_address}?`);
            card.setButton({ text: 'Explore Remote Working', url: "https://pwcuk.cobrainer.com/create-job-posting" });
            agent.add(card);
            agent.add(new Suggestion('Book trip'));
            agent.add(new Suggestion('Request formal assignment'));
          }
          break;

        default:
          break;
      }



      // error
    } catch (error) {
      console.log(error);
    }
  }

  async function booktrip(agent) {
    try {

      let originAddress = "London";
      let targetAddress = request.body.queryResult.outputContexts[0].parameters["geo-city"];
      let periodtDate = request.body.queryResult.outputContexts[0].parameters["date-period"];
      let date1 = periodtDate.startDate;
      let date2 = periodtDate.endDate;
      let diff = Math.floor((Date.parse(date2) - Date.parse(date1)) / 86400000);
      let job_roles = request.body.queryResult.parameters.job_roles;
      console.log(`agent.requestSource: ${agent.TELEGRAM} , job_roles: ${job_roles} `);
      var values = await get_price(originAddress, targetAddress);
      let distance = values.distance;
      let Co2_emission = Math.round(distance * 0.285);
      let target_country = values.target_country;
      // let price = Math.round(values.price * 0.0140887);
      let price = '150';
      // var price = values.price;
      // let flight_time = values.flight_time;

      let homeperc = await job_search(job_roles);
      console.log(`homeperc: ${homeperc} `);

      agent.add(new Suggestion('personally undertaking an internal audit'));
      agent.add(new Suggestion('client meetings'));
      agent.add(new Suggestion('internal meeting'));
      agent.add(new Suggestion('attending on-the-job training'));
      agent.add(new Suggestion('other'));
      agent.add("Please confirm if you will be carrying out one of the the following activities");
      let ctx = { 'name': 'transfer_params', 'lifespan': 4, 'parameters': { 'homeperc': homeperc, 'job_roles': job_roles, 'target_country': target_country, 'price': price, 'diff': diff, 'Co2_emission': Co2_emission } };
      agent.context.set(ctx);


    } catch (error) {
      console.log(error);
    }
  }

  async function makeBooktrip(agent) {
    console.log("makeBooktrip");
    var startDate;
    var target_address;
    if (typeof request.body.queryResult.outputContexts[1].parameters["date-period"] == 'undefined') {
      startDate = request.body.queryResult.outputContexts[0].parameters["date-period"].startDate;
      target_address = request.body.queryResult.outputContexts[0].parameters["geo-city"];
    } else {
      startDate = request.body.queryResult.outputContexts[1].parameters["date-period"].startDate;
      target_address = request.body.queryResult.outputContexts[1].parameters["geo-city"];
    }

    const dateTimeStart = new Date(Date.parse(startDate.split('T')[0] + 'T' + startDate.split('T')[1].split('+')[0] + timeZoneOffset));
    const dateTimeEnd = new Date(new Date(dateTimeStart).setHours(dateTimeStart.getHours() + 1));
    console.log(`starttime: ${dateTimeStart}, endtime: ${dateTimeEnd}`);
    const appointmentTimeString = dateTimeStart.toLocaleString(
      'en-US',
      { month: 'long', day: 'numeric', hour: 'numeric', timeZone: timeZone }
    );

    // Check the availibility of the time, and make an appointment if there is time on the calendar
    return createCalendarEvent(dateTimeStart, dateTimeEnd, target_address).then(() => {
      agent.add(`Done. I will send you completed social security and PWD certificates shortly. I will also arrange a health screening appointment for an available slot in your calendar. Have a great trip to ${target_address}.`);
    }).catch(() => {
      agent.add(`you have already used that timeslot.`);
    });
  }

  function fallback(agent) {
    agent.add(`I'm sorry, can you try again?`);
  }

  let intentMap = new Map();

  intentMap.set('Default Fallback Intent', fallback);
  intentMap.set('book.trip', booktrip);
  intentMap.set('book.trip-activity', booktrip_activity);
  intentMap.set('book.trip-activity - custom', makeBooktrip);
  agent.handleRequest(intentMap);
});

async function get_price(origin, target) {
  return new Promise((resolve, reject) => {
    try {
      console.log(`origin: ${origin}`);
      console.log(`target: ${target}`);
      var options = {
        'method': 'GET',
        'url': 'https://www.distance24.org/route.json?stops=' + origin + '|' + target,
        'headers': {
        }
      };
      console.log(options.url);
      requests(options, async function (error, response) {
        if (error) {
          reject(error);
        }
        console.log(`response.body: ${response.body}`);
        var body = JSON.parse(response.body);
        var distance = body.distance;
        var countrycode = body.stops[1].countryCode;
        var target_country = countries[countrycode].name;

        // getting IATA code of origin address
        var origin_latitude = body.stops[0].latitude;
        var origin_longitude = body.stops[0].longitude;
        var origin_IATA_url = 'http://iatageo.com/getCode/' + origin_latitude + '/' + origin_longitude;
        var options_origin = {
          'method': 'GET',
          'url': origin_IATA_url,
          'headers': {
          }
        };
        console.log(`res_origin url for getting IATA code ${options_origin.url}`);
        var res_origin = await rp(options_origin);
        console.log(`res_origin ${res_origin}`);
        var origin_IATA_code = JSON.parse(res_origin).IATA;
        // var origin_IATA_code = 'LHR';

        // getting IATA code of target address
        var target_latitude = body.stops[1].latitude;
        var target_longitude = body.stops[1].longitude;
        var target_IATA_url = 'http://iatageo.com/getCode/' + target_latitude + '/' + target_longitude;
        var options_target = {
          'method': 'GET',
          'url': target_IATA_url,
          'headers': {
          }
        };
        console.log(`res_target url for getting IATA code ${options_target.url}`);
        var res_target = await rp(options_target);
        console.log(`res_target ${res_target}`);
        var target_IATA_code = JSON.parse(res_target).IATA;


        /*
                //getting flight ticket price
                var options = {
                  'method': 'GET',
                  'url': `https://travelpayouts-travelpayouts-flight-data-v1.p.rapidapi.com//v1/prices/direct/?destination=${target_IATA_code}&origin=${origin_IATA_code}`,
                  'headers': {
                    'x-rapidapi-host': 'travelpayouts-travelpayouts-flight-data-v1.p.rapidapi.com',
                    'x-rapidapi-key': apikey,
                    'X-Access-Token': 'cdabee9c1dee387ac09e5f14c5e95a05'
                  }
                };
        
                var res_price = await rp(options).catch(error => {
                  console.log(error);
                });
                console.log(`price: ${res_price}`);
                var code = JSON.stringify(JSON.parse(res_price).data).slice(2, 5);
                // console.log(`price: ${JSON.parse(res_price).data[code]["0"].price}`);
                var price = JSON.parse(res_price).data[code]["0"].price;
                console.log(`price: ${price}`);
        */

        var prop_para = { distance: distance, /*price: price, */ target_country: target_country };
        resolve(prop_para);

      });
    } catch (error) {
      console.log(error);
    }



  });

}

async function get_carbonemission(distance) {
  return new Promise((resolve, reject) => {
    try {
      var options = {
        method: 'GET',
        url: 'https://carbonfootprint1.p.rapidapi.com/CarbonFootprintFromFlight',
        qs: { distance: distance, type: 'DomesticFlight' },
        headers: {
          'x-rapidapi-host': 'carbonfootprint1.p.rapidapi.com',
          'x-rapidapi-key': apikey,
          useQueryString: true
        }
      };

      requests(options, function (error, response, body) {
        if (error) throw new Error(error);

        console.log(body);
        var carbon = JSON.parse(body).carbonEquivalent + 'kg';
        console.log(`Co2 emission: ${carbon}`);
        resolve(carbon);
      });
    } catch (error) {
      console.log(error);
      reject(error);
    }

  });
}

async function need_pwd(activity, target_country, diff) {
  return new Promise((resolve, reject) => {
    var workbook = XLSX.readFile('Chatbot materials.xlsx');
    var sheet_name_list = workbook.SheetNames;
    var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    var count = 0;
    var n_pwd;
    xlData.forEach(row => {
      count++;
      if (row.__EMPTY == target_country) {
        console.log(`____________${activity}___________`);
        switch (activity) {
          case 'Personally undertaking an internal audit':
            let period1 = row.__EMPTY_2;
            console.log(`____________${period1}___________`);
            if (period1 <= diff) {
              n_pwd = true;
              resolve({ immi: row.__EMPTY_1, n_pwd: n_pwd, period: period1 });

            } else {
              n_pwd = false;
              resolve({ immi: row.__EMPTY_1, n_pwd: n_pwd, period: period1 });
              console.log("false");

            }
            break;
          case 'Client meetings':
            let period2 = row.__EMPTY_4;
            console.log(`____________${period2}___________`);
            if (period2 <= diff) {
              n_pwd = true;
              resolve({ immi: row.__EMPTY_3, n_pwd: n_pwd, period: period2 });
              console.log("test");

            } else {
              n_pwd = false;
              console.log(`n_pwd: ${n_pwd}`);

              resolve({ immi: row.__EMPTY_3, n_pwd: n_pwd, period: period2 });

            }
            break;
          case 'Internal meeting':
            let period3 = row.__EMPTY_6;
            console.log(`____________${period3}___________`);
            if (period3 <= diff) {
              n_pwd = true;
              resolve({ immi: row.__EMPTY_5, n_pwd: n_pwd, period: period3 });
              console.log("test");

            } else {
              n_pwd = false;
              resolve({ immi: row.__EMPTY_5, n_pwd: n_pwd, period: period3 });

            }
            break;
          case 'Attending on-the-job training':
            let period4 = row.__EMPTY_8;
            console.log
              (`____________${period4}___________`);
            if (period4 <= diff) {
              n_pwd = true;
              resolve({ immi: row.__EMPTY_7, n_pwd: n_pwd, period: period4 });
              console.log("test");

            } else {
              n_pwd = false;
              resolve({ immi: row.__EMPTY_7, n_pwd: n_pwd, period: period4 });

            }
            break;
          case 'Other':

            resolve({ immi: 'NP', n_pwd: true, period: '0' });

            break;
          default:
            break;
        }
      }
      else if (count == xlData.length && typeof n_pwd == "undefined") {
        console.log(typeof n_pwd);
        console.log(n_pwd);
        resolve({ immi: 'NP', n_pwd: true, period: '0' });
      }

      return true;

    });
  });

}

async function job_search(job_roles) {
  return new Promise(async (resolve, reject) => {
    try {
      const username = 'center_internal'
      const password = '9941hst'
      const onet_ws = new OnetWebService(username, password)
      console.log(`onet_ws: ${JSON.stringify(onet_ws)}`);
      const kwquery = job_roles;
      const kwresults = await onet_ws.call('online/search', {
        keyword: kwquery,
        end: 3
      })
      check_for_error(kwresults)
      if (!kwresults.hasOwnProperty('occupation') || !kwresults.occupation.length) {
        console.log('No relevant occupations were found.')
        console.log('')
      } else {
        console.log('Most relevant occupations for "' + kwquery + '":')
        var workbook = XLSX.readFile('Physical_Proximity.xls');
        var sheet_name_list = workbook.SheetNames;
        var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
        var sum = 0;

        for (let occ of kwresults.occupation) {
          console.log('  ' + occ.code + ' - ' + occ.title)
          xlData.forEach(record => {
            // console.log(`record: ${JSON.stringify(record)}`);
            if (record.__EMPTY == occ.code) {
              sum = sum + record["Browse by O*NET Data"];
              console.log(`value: ${sum}`);
            }
          });
        }
        console.log(`percent is ${sum / 3}`);
        if (sum / 3 >= 50) {
          resolve("yes");
        } else {
          resolve("no");
        }


      }
    } catch (error) {
      console.error(error.message)
      reject("error");
    }
  });
}

function check_for_error(service_result) {
  if (service_result.hasOwnProperty('error')) {
    throw new Error(service_result.error)
  }
}

function createCalendarEvent(dateTimeStart, dateTimeEnd, appointment_type) {
  return new Promise((resolve, reject) => {
    calendar.events.list({
      auth: serviceAccountAuth, // List events for time period
      calendarId: calendarId,
      timeMin: dateTimeStart.toISOString(),
      timeMax: dateTimeEnd.toISOString()
    }, (err, calendarResponse) => {
      console.log(`response: ${calendarResponse.data.items}`);

      // Check if there is a event already on the Calendar
      if (err || calendarResponse.data.items.length > 0) {
        console.log(`error: ${err}    item: ${calendarResponse.data.items.length}`)
        reject(err || new Error('Requested time conflicts with another appointment'));
      } else {
        // Create event for the requested time period
        calendar.events.insert({
          auth: serviceAccountAuth,
          calendarId: calendarId,
          resource: {
            summary: 'Trip book to ' + appointment_type, description: 'travel to ' + appointment_type,
            start: { dateTime: dateTimeStart },
            end: { dateTime: dateTimeEnd }
          }
        }, (err, event) => {
          err ? reject(err) : resolve(event);
        }
        );
      }
    });
  });
}

