export class cachingService {
    private ttl:number = 0;

    constructor(ttl:number) {
        this.ttl=ttl;
    }

    public setWithGlobalExpiry(key, value) {
        const now = new Date();
        let ttl:number = this.ttl;
        // `item` is an object which contains the original value
        // as well as the time when it's supposed to expire
        const item = {
            value: value,
            expiry: now.getTime() + ttl,
        };
        //try and clear cache if full
        try{
            localStorage.setItem(key, JSON.stringify(item));
        }
        catch{
            //clear all cache if full
            this.clearAllCache();
        }
    }

    public setWithExpiry(key, value, ttl) {
        const now = new Date();
        // `item` is an object which contains the original value
        // as well as the time when it's supposed to expire
        const item = {
            value: value,
            expiry: now.getTime() + ttl,
        };
        //try and clear cache if full
        try{
            localStorage.setItem(key, JSON.stringify(item));
        }
        catch{
            //clear all cache if full
            this.clearAllCache();
        }
    }

    public getWithExpiry(key) {
        const itemStr = localStorage.getItem(key);
        // if the item doesn't exist, return null
        if (!itemStr) {
            return null;
        }
        const item = JSON.parse(itemStr);
        const now = new Date();
        // compare the expiry time of the item with the current time
        if (now.getTime() > item.expiry) {
            // If the item is expired, delete the item from storage
            // and return null
            localStorage.removeItem(key);
            return null;
        }
        return item.value;
    }

    public removeCache(key){
        localStorage.removeItem(key);
    }

    private clearAllCache(){
        //empty cache for CDBApps
        console.log("clearing local storage for CDB Apps");
        var key; 
        for (var i = 0; i < localStorage.length; i++) 
        { 
            key = localStorage.key(i); 
            if(key.indexOf("CDB")>-1)
            { 
                localStorage.removeItem(key); 
            } 
        }
        console.log("finished clearing local storage");
    }
}
