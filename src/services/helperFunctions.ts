export class helperFunctions {

    constructor() {
        
    }

    public getTimeNow():string{
        var today = new Date();
        let mins:string=today.getMinutes().toString();
        let sec:string=today.getSeconds().toString();
        if(today.getSeconds()<10){
          sec="0"+today.getSeconds().toString();
        }
        if(today.getMinutes()<10){
          mins="0"+today.getMinutes().toString();
        }
        if(today.getMinutes())
        return today.getHours() + ":" + mins + ":" + sec;
      }




    public reportDebug(logmessage:string):void{
        let debugswitch:string = this.getParameterByName('cdbdebug');
        if(debugswitch){
            console.log(logmessage);
        }
    }

    private getParameterByName(name):string {
        let url:string = window.location.href;
        name = name.replace(/[\[\]]/g, "\\$&");
        var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
            results = regex.exec(url);
        if (!results) return null;
        if (!results[2]) return '';
        return decodeURIComponent(results[2].replace(/\+/g, " "));
    }


}
