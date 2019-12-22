declare module "google" {
    interface IGoogle {
        charts: any;
        visualization: any;
    }
 
    var google: IGoogle;
    export = google;
}