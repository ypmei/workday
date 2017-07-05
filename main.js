'use strict';
var xlsx = require('node-xlsx')
var fs = require('fs')
var _ = require('lodash')
var moment = require('moment')

const MODE = 1; //1是两两比差，2是不相交的差

var workbook  = xlsx.parse(`${__dirname}/xlsx/dingding.xlsx`)
var worksheet = workbook[0].data

var infoData = _.reject(worksheet, (sheet)=>(sheet.length === 23&&sheet[0]!=='姓名'))
var listData = _.filter(worksheet, (sheet)=>(sheet.length === 23&&sheet[0]!=='姓名'))

var u8book  = xlsx.parse(`${__dirname}/xlsx/u8.xlsx`)
var u8sheet = u8book[0].data

var u8Data = _.filter(u8sheet, (s)=>(s[6]!=='工时'))

var format = {
  timeVal:(v)=>{
    return moment(1503504000000+v).format('HH:mm:ss')
  },
  hourVal:(v)=>{
    let hour = Math.floor(v/1000/60/60).toString()
    let mmss = moment(1501516800000+v).format('mm:ss')
    return hour+':'+mmss
  },
  otVal:(v)=>{
    let offset = v-8*60*60*1000
    return offset > 0 ? offset : 0
  },
  sliceToPair:(data)=>{
    const len = data.length
    let res = []
    for(var i=0; i < len; i+=2){
       res.push(data.slice(i,i+2));
    }
    return res
  },
  sliceToBoth:(data)=>{
    const len = data.length
    let res = []
    for(var i=0; i < len-1; i++){
       res.push(data.slice(i,i+2));
    }
    return res
  }
}

function parseTime(data){
  let rows = []
  const groupData = _.groupBy(data,(row)=>row[5]) //按时间分组
  _.each(groupData, (val,key)=>{
    if(val.length === 1) {
      rows.push(val[0])
    }else{
      const pair = MODE === 1 ? format.sliceToBoth(val) : format.sliceToPair(val)
      pair.forEach((p,ind)=>{
        const isOdd = p.length === 1
        if(isOdd){
          p.unshift(pair[ind-1][1])//如果是奇数，取最后两个
        }
        const [one, two] = p
        const timeOne = moment(`${one[5]} ${one[6]}`).valueOf()
        const timeTwo = moment(`${two[5]} ${two[6]}`).valueOf()
        const timeDiff = timeTwo-timeOne
        const timeValue = format.timeVal(timeDiff)
        two[17] = timeDiff
        two[18] = timeValue

        let weekday = moment(timeTwo).format('E')
        if([1,2,3,4,5].includes(parseInt(weekday))){
          let overtime = format.otVal(timeDiff)
          two[23] = overtime
          two[24] = overtime ? format.timeVal(overtime) : null
        }
        if([6,7].includes(parseInt(weekday))){
          two[27] = timeDiff
          two[28] = timeValue
        }
        if(MODE === 1){
          //[1,2,3,4] = [1,2],[2,3],[3,4]
          if(ind === 0){
            rows.push(one)
          }
          rows.push(two)
        }else{
          if(!isOdd){
            rows.push(one)
          }
          rows.push(two)
        }
      })
    }
  })
  return rows
}

function countTime(data,type){
  let result = []
  let group = []
  if(type==='day'){
    group = _.groupBy(data,(row)=>row[5])
  }else{
    group = _.groupBy(data,(row)=>{
      let timestamp = moment(`${row[5]} 00:00:00`).valueOf()
      return moment(timestamp).format('M')
    })
  }
  _.each(group, (g,k)=>{
    const dayTotal = _.chain(g).map((row)=>parseInt(row[17])).filter().reduce((x,y)=>x+y).value()
    const overtimeWeek = _.chain(g).map((row)=>parseInt(row[23])).filter().reduce((x,y)=>x+y).value()
    const overtimeWeekend = _.chain(g).map((row)=>parseInt(row[27])).filter().reduce((x,y)=>x+y).value()

    if(dayTotal){
      if(type==='day'){
        g[g.length-1][19] = dayTotal
        g[g.length-1][20] = format.hourVal(dayTotal)
      }else{
        g[g.length-1][21] = dayTotal
        g[g.length-1][22] = format.hourVal(dayTotal)

        if(overtimeWeek){
          g[g.length-1][25] = overtimeWeek
          g[g.length-1][26] = format.hourVal(overtimeWeek)
        }
        if(overtimeWeekend){
          g[g.length-1][29] = overtimeWeekend
          g[g.length-1][30] = format.hourVal(overtimeWeekend)
        }
      }
    }
    result = result.concat(g)
  })
  return result
}

function appendOrder(row){
  const name = row[0].replace(/\[离职\]/,'')
  let timestamp = moment(`${row[5]} ${row[6]}`).valueOf()
  timestamp = moment(timestamp).format('YYYY-MM-DD HH:mm')
  timestamp = moment(timestamp+':00').valueOf()

  let orderRow = _.chain(u8Data).filter((uRow)=>uRow[3]===name).filter((v)=>{
    const end = moment(v[4]).valueOf()
    const start = moment(v[5]).valueOf()
    return timestamp >= start && timestamp < end
  }).first().value()

  if(orderRow){
    row[14] = orderRow[0]
    row[15] = orderRow[1]
    row[16] = orderRow[2]
  }
  return row
}

var sheets = _.chain(listData).map((row)=>{
    const newRow = appendOrder(row)
    return row
  }).groupBy((list)=>list[0]).map((val,key)=>{
    let data = parseTime(val)
    data = countTime(data,'day')
    data = countTime(data,'month')

    return {
      name: key,
      data: infoData.concat(data)
    }
  }).value()

var buffer = xlsx.build(sheets)
fs.writeFileSync('data/dingding.xlsx', buffer, {'flag':'w'})
