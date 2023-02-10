// var encodedWebUrl=encodeURIComponent(`${this.BaseURL}/ChannelDeepLink.html&label=DeepLink`);
// GetDeepLinkTabChannel = (subEntityId: any, ID: any, Desc: any, channelId: any,AppID: string,EntityID: string)=>{

//    let taskContext = encodeURIComponent(`{"subEntityId": "${subEntityId}","channelId":"${channelId}"}`);
//      return {
//       linkUrl:"https://teams.microsoft.com/l/entity/"+AppID+"/"+EntityID+"?webUrl=" + encodedWebUrl + "&context=" + taskContext,
//       ID:ID,
//       TaskText:Desc
//      }
//    }

export const getDeepLinkTabStatic = (subEntityId: string, ID: string, Desc: string, AppID: string | undefined)=>{
   let taskContext = encodeURI(`{"subEntityId": "${subEntityId}"}`);
     return {
      linkUrl:"https://teams.microsoft.com/l/entity/"+AppID+"/"+process.env.Tab_Entity_Id +"?context=" + taskContext,
      ID:ID,
      TaskText:Desc
     }    
};

// module.exports= {
//    // GetDeepLinkTabChannel,
//    GetDeepLinkTabStatic
// }
