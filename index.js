'use strict';
var xlsx = require('node-xlsx')
var fs = require('fs')
var _ = require('lodash')
var moment = require('moment')

var workbook  = xlsx.parse(`${__dirname}/xlsx/dingding.xlsx`)
var worksheet = workbook[0].data

var infoData = _.reject(worksheet, (sheet)=>(sheet.length === 23&&sheet[0]!=='姓名'))
var listData = _.filter(worksheet, (sheet)=>(sheet.length === 23&&sheet[0]!=='姓名'))


var format = {
  timeVal:(v)=>{
    return moment(1503504000000+v).format('HH:mm:ss')
  },
  hourVal:(v)=>{
    let hour = Math.floor(v/1000/60/60).toString()
    let mmss = moment(1501516800000+v).format('mm:ss')
    return hour+':'+mmss
  },
  sliceToPair:(data)=>{
    const len = data.length
    let res = []
    for(var i=0; i < len; i+=2){
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
      const pair = format.sliceToPair(val)
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
        two[14] = timeDiff
        two[15] = timeValue
        if(!isOdd){
          rows.push(one)
        }
        rows.push(two)
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
    // console.log(group)
  }
  _.each(group, (g,k)=>{
    const dayTotal = _.chain(g).map((row)=>parseInt(row[14])).filter().reduce((x,y)=>x+y).value()
    if(dayTotal){
      if(type==='day'){
        g[g.length-1][16] = dayTotal
        g[g.length-1][17] = format.hourVal(dayTotal)
      }else{
        g[g.length-1][18] = dayTotal
        g[g.length-1][19] = format.hourVal(dayTotal)
      }
    }
    result = result.concat(g)
  })
  return result
}

var sheets = _.chain(listData).groupBy((list)=>list[0]).map((val,key)=>{
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
