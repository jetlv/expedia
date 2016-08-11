/// <reference path="./include.d.ts" />
var request = require('request');
var fs = require('fs');
var async = require('async');
var ew = require('node-xlsx');
var cheerio = require('cheerio');

/** define file name */
const filePath = "expedia.xlsx";
/**
 * excel prepration
 */
var sheet;
if (fs.existsSync(filePath)) {
    fs.createReadStream(filePath).pipe(fs.createWriteStream('backup.xlsx'));
    sheet = ew.parse(fs.readFileSync(filePath))[0];
} else {
    var columns = ["DATE_EXTRACT", "HOTEL_ID", "HOTEL_NAME", "H_EXPRAT", "H_CAT", "H_LOC", "ROOMTYPE_ID", "ROOMTYPE", "RATEPLAN", "RATE_CAT", "RATE_NAME", "BEDTYPE", "ROOM_SIZE", "RATE_T0", "RATE_T7", "RATE_T14", "RATE_T28", "RATE_T56", "RATE_T102"];
    sheet = { name: 'result', data: [] };
    sheet.data.push(columns);
}

var rows = sheet.data;


function composeHar(hotelId, chkin, chkout, cookie, token) {
    var harTemp = {
        "method": "GET",
        "url": '',
        "httpVersion": "HTTP/1.1",

        "headers": [
            {
                "name": "Accept",
                "value": "application/json, text/javascript, */*; q=0.01"
            },
            {
                "name": "Accept-Encoding",
                "value": "gzip, deflate, br"
            },
            {
                "name": "Accept-Language",
                "value": "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3"
            },
            {
                "name": "Connection",
                "value": "keep-alive"
            },
            {
                "name": "Host",
                "value": "www.expedia.com.hk"
            },
            {
                "name": "User-Agent",
                "value": "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:47.0) Gecko/20100101 Firefox/47.0"
            },
            {
                "name": "X-Requested-With",
                "value": "XMLHttpRequest"
            }
        ],
        "queryString": [
            {
                "name": "adults",
                "value": "1"
            },
            {
                "name": "children",
                "value": "0"
            },
            {
                "name": "isVip",
                "value": "false"
            },
            {
                "name": "ts",
                "value": "1470703188974"
            }
        ],
    }
    var hkhar = harTemp;

    var start = {
        "name": "chkin",
        "value": chkin
    };
    var end =
        {
            "name": "chkout",
            "value": chkout
        };

    var ck = {
        "name": "Cookie",
        "value": cookie
    }

    var tk = {
        "name": "token",
        "value": token
    }
    hkhar.queryString.push(start);
    hkhar.queryString.push(end);
    hkhar.queryString.push(token);
    hkhar.url = "https://www.expedia.com.hk/api/infosite/" + hotelId + "/getOffers?token=" + token + "&chkin=" + chkin + "&chkout=" + chkout + "&adults=1&children=0";
    hkhar.headers.push(ck);

    return hkhar;
}
var __getSessions = function (resp) {
    var cookies = [];
    var fullArr = resp.headers['set-cookie'];
    for (var i in fullArr) {
        cookies.push(fullArr[i].split(';')[0]);
    }

    return cookies.join("; ");
}

var hotels = [
    {
        "name": "Altira Macau",
        "id": "10091860",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Altira-Macau.h10091860.Hotel-Information"
    },
    {
        "name": "Banyan Tree Macau",
        "id": "4282350",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Banyan-Tree-Macau.h4282350.Hotel-Information"
    },
    {
        "name": "Broadway Macau",
        "id": "10106413",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Broadway-Macau.h10106413.Hotel-Information"
    },
    {
        "name": "City of Dreams-Crown Towers Macau",
        "id": "2871315",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-City-Of-Dreams-Crown-Towers-Macau.h2871315"
    },
    {
        "name": "Conrad Macao Cotai Central",
        "id": "4944702",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Conrad-Macao-Cotai-Central.h4944702.Hotel-Information"
    },
    {
        "name": "Four Seasons Macao at Cotai Strip",
        "id": "2363104",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Four-Seasons-Hotel-Macao-At-Cotai-Strip.h2363104"
    },
    {
        "name": "Galaxy Macau",
        "id": "4359010",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Galaxy-Macau.h4359010.Hotel-Information"
    },
    {
        "name": "Grand Hyatt Macau",
        "id": "2844548",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Grand-Hyatt-Macau.h2844548.Hotel-Information"
    },
    {
        "name": "Grand Lisboa",
        "id": "2867646",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Grand-Lisboa-Macau.h2867646.Hotel-Information"
    },
    {
        "name": "Grandview Hotel Macau",
        "id": "1042400",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Grandview-Hotel-Macau.h1042400.Hotel-Information"
    },
    {
        "name": "Hard Rock Hotel",
        "id": "2759501",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Hard-Rock-Hotel.h2759501.Hotel-Information"
    },
    {
        "name": "Holiday Inn Macau Cotai Central",
        "id": "4894016",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Holiday-Inn-Macao-Cotai-Central.h4894016"
    },
    {
        "name": "Hotel Okura Macau",
        "id": "4346833",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Hotel-Okura-Macau.h4346833"
    },
    {
        "name": "JW Marriott Hotel Macau",
        "id": "10224807",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-JW-Marriott-Hotel-Macau.h10224807.Hotel-Information"
    },
    {
        "name": "MGM Macau",
        "id": "1795541",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-MGM-MACAU.h1795541.Hotel-Information"
    },
    // {
    //     "name" : "MGM Cotai" //not open yet
    // },
    // {
    //     "name" : "Parisian Macao" //not open yet
    // },
    {
        "name": "Regency Hotel Macau",
        "id": "21481",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Regency-Hotel-Macau.h21481.Hotel-Information"
    },
    {
        "name": "Studio City",
        "id": "12701563",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Studio-City.h12701563.Hotel-Information"
    },
    {
        "name": "The Ritz Carlton Macau",
        "id": "10043106",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-The-Ritz-Carlton-Macau.h10043106.Hotel-Information"
    },
    {
        "name": "St. Regis Macao",
        "id": "11927099",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-The-St-Regis-Macao.h11927099.Hotel-Information"
    },
    {
        "name": "The Venetian Macao Resort",
        "id": "1691530",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-The-Venetian-Macao-Resort.h1691530.Hotel-Information"
    },
    {
        "name": "Wynn Macau",
        "id": "1503945",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Wynn-Macau.h1503945.Hotel-Information"
    },
    {
        "name": "Wynn Palace", //not open yet
        "id": "15935917",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Wynn-Palace.h15935917"
    }
];


function fetchRate(ckin, ckout, hotel, outerCallback) {
    var hotelUrl = hotel.baseUrl;
    var hotelId = hotel.id;
    request({ url: hotelUrl, method: 'GET', gzip: true }, function (err, resp, hbody) {
        if (!resp) {
            var badRow = [];
            badRow.push(new Date(ckin + ' 08:00:00'));
            badRow.push(hotel.id);
            badRow.push(hotel.name);
            badRow.push("Unexpected error, no response : " + err ? err : "");
            rows.push(badRow);
            outerCallback();
            return;
        }
        var cookie = __getSessions(resp);
        var token = hbody.match(/infosite\.token =.*/g)[0].match(/=.*/g)[0].replace(/[=\;']/g, '').trim();

        var entities = [];
        var offsets = [0, 7, 14, 28, 56, 102];
        async.mapLimit(offsets, 1, function (offset, callback) {
            var thar = composeHar(hotelId, urlComposer(offset).chkin, urlComposer(offset).chkout, cookie, token);
            var offsetEntity;
            request({ har: thar, gzip: true }, function (err, resp, ibody) {
                try {
                    var iOffers = JSON.parse(ibody).offers;
                } catch (err) {
                    setTimeout(function () {
                        offsetEntity = {
                            "offset": offset,
                            "offers": []
                        }
                        entities.push(offsetEntity);
                        callback();
                    }, 7000);
                    return;
                }
                if (!iOffers) {
                    offsetEntity = {
                        "offset": offset,
                        "offers": []
                    }
                    entities.push(offsetEntity);
                    setTimeout(function () {
                        callback();
                    }, 7000);
                    return;
                }
                offsetEntity = {
                    "offset": offset,
                    "offers": iOffers
                }
                entities.push(offsetEntity);
                setTimeout(function () {
                    console.log('offset ' + offset + ' was fetched');
                    callback();
                }, 2000);
            });
        }, function (err) {
            if (err) console.log(err);
            // fs.writeFileSync('entities.json', JSON.stringify(entities));
            var keys = [];
            var singleRows = [];
            entities.forEach(function (offsetEntity, index, array) {
                offsetEntity.offers.forEach(function (entity, index, array) {
                    var offset = offsetEntity.offset;
                    var key = getKeyOfOffer(entity);
                    console.log(offset + ' - ' + key);
                    if (keys.indexOf(key) == -1) {
                        var date = new Date(ckin + ' 08:00:00');
                        var hotelID = parseInt(hotel.id);
                        var hotelName = hotel.name;
                        var $ = cheerio.load(hbody);
                        var hotelRate = parseInt($('.rating-number').text());
                        var hcat = ($('#license-plate .visuallyhidden').text().match(/[\d\.]+/)[0] + ' Stars').replace(/\.0/g, '');
                        var hloc = $('.street-address').eq(0).text() + ', ' + $('.city').eq(0).text();
                        var roomTypeCode = parseInt(entity.roomTypeCode);
                        var ratePlanCode = parseInt(entity.ratePlanCode);
                        var rateCatArray = [];
                        for (var cat in entity.amenities) {
                            rateCatArray.push(entity.amenities[cat]);
                        }
                        var rateCat = rateCatArray.join(',');
                        var rateName = entity.freeCancellable ? "REFUND" : "NONREF";
                        var roomName = entity.roomName;
                        var rateStr = hbody.match('\\{"rooms":\\[.*"ratePlans":.*\\}')[0];
                        var rateAndPlan = JSON.parse(rateStr);
                        var bedtype = '';
                        var roomSize = '';
                        rateAndPlan.rooms.forEach(function (item, index, array) {
                            if (item.roomTypeCode === roomTypeCode) {
                                bedtype = item.beddingOptions.join(',').trim();
                                roomSize = item.roomSquareMeters;
                            }
                        });
                        var t = '';
                        if (!entity.price) {
                            t = 'N/A';
                        } else {
                            t = entity.price.displayPrice;
                        }
                        var row = [];
                        row.push(date);
                        row.push(hotelID);
                        row.push(hotelName);
                        row.push(hotelRate);
                        row.push(hcat);
                        row.push(hloc);
                        row.push(roomTypeCode);
                        row.push(roomName);
                        row.push(ratePlanCode);
                        row.push(rateCat);
                        row.push(rateName);
                        row.push(bedtype);
                        row.push(roomSize);
                        row[13 + offsets.indexOf(offset)] = t;
                        [0,1,2,3,4,5].forEach(function(v, i, a) {
                            //set 'N/A' as default value, aviod empty cell
                            row.push('N/A');
                        });
                        keys.push(key);
                        singleRows.push({
                            "key": key,
                            "rs": row
                        });
                    } else {
                        singleRows.forEach(function (sr, index, array) {
                            if (sr.key === key) {
                                var t = '';
                                if (!entity.price) {
                                    t = 'N/A';
                                } else {
                                    t = entity.price.displayPrice;
                                }
                                sr.rs[13 + offsets.indexOf(offset)] = t;
                            }
                        });
                    }
                });
            });

            singleRows.forEach(function (item, index, value) {
                rows.push(item.rs);
            });

            setTimeout(function () {
                console.log(hotel.name + ' was done');
                var buffer = ew.build([sheet]);
                fs.writeFileSync(filePath, buffer);
                outerCallback();
            }, 2000);
        });



        function getKeyOfOffer(offer) {
            return offer.hotelID + offer.roomTypeCode + offer.ratePlanCode + (offer.freeCancellable ? "refund" : "nonref");
        }
    });
}


/**
 * Date extension
 */
Date.prototype.Format = function (fmt) {
    var o = {
        "M+": this.getMonth() + 1,
        "d+": this.getDate(),
        "h+": this.getHours(),
        "m+": this.getMinutes(),
        "s+": this.getSeconds(),
        "q+": Math.floor((this.getMonth() + 3) / 3),
        "S": this.getMilliseconds()
    };
    if (/(y+)/.test(fmt))
        fmt = fmt.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
    for (var k in o)
        if (new RegExp("(" + k + ")").test(fmt))
            fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
    return fmt;
}

function urlComposer(choosenOffset) {
    var params = '#adults=1&children=0';
    var ckinDate = new Date();
    ckinDate.setDate(new Date().getDate() + choosenOffset);
    var ckin = ckinDate.Format('yyyy/MM/dd');
    var ckoutDate = new Date();
    ckoutDate.setDate(new Date().getDate() + choosenOffset + 1);
    var ckout = ckoutDate.Format('yyyy/MM/dd');
    params += '&chkin=' + ckin;
    params += '&chkout=' + ckout;

    return { params: params, chkin: ckin, chkout: ckout };
}


function run() {
    console.log('Task started ' + new Date())
    var chkin = urlComposer(0).chkin;
    var chkout = urlComposer(0).chkout;
    var test = [{
        "name": "Altira Macau",
        "id": "10091860",
        "baseUrl": "https://www.expedia.com.hk/en/Macau-Hotels-Altira-Macau.h10091860.Hotel-Information"
    }];
    async.mapLimit(/** hotels */ test, 1, function (hotel, callback) {
        fetchRate(chkin, chkout, hotel, callback);
    }, function (err) {
        if (err) console.log(err);
        var buffer = ew.build([sheet]);
        fs.writeFileSync(filePath, buffer);
        console.log('Everything was done');
    });
}

run();
// setInterval(run, 1000 * 60 * 60 * 24);
