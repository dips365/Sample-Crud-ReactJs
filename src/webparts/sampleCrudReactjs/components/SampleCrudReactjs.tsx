import * as React from 'react';
import styles from './SampleCrudReactjs.module.scss';
import { ISampleCrudReactjsProps } from './ISampleCrudReactjsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ISampleCrudReactjsStates } from "./ISampleCrudReactjsStates";
import { Environment,EnvironmentType } from "@microsoft/sp-core-library";
import { SPHttpClient,SPHttpClientResponse } from "@microsoft/sp-http";
import { IListItem } from './IListItem';

export default class SampleCrudReactjs extends React.Component<ISampleCrudReactjsProps, ISampleCrudReactjsStates> {


  constructor(props:ISampleCrudReactjsProps,states:ISampleCrudReactjsStates){
    super(props);
    this.state={
      status:"Yeahhhh  Ready to go",
      items:[]
    };
  }

  public render(): React.ReactElement<ISampleCrudReactjsProps> {

    const items:JSX.Element[] = this.state.items.map((item:IListItem,i:number):JSX.Element=>{
      return(
        <li>{item.Title} ({item.ID})</li>
      );
    });
    return (
      <div className={ styles.sampleCrudReactjs }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint FrameWork!</span>
              <p className={ styles.subTitle }>CRUD operations using React javascript Library.</p>
              {/* <p className={ styles.description }>{escape(this.props.description)}</p> */}
              {/* <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a> */}
              <p className={styles.description}>List Name: {escape(this.props.ListName)}</p>

              {/*Add buttons for CRUD Operations*/}

              <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
                <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
                  <a href="#" className={`${styles.button}`} onClick={() => this.AddItem()}>
                    <span className={styles.label}>Create item</span>
                  </a>&nbsp;
                  <a href="#" className={`${styles.button}`} onClick={() => this.GetItems()}>
                    <span className={styles.label}>Read item</span>
                  </a>
                </div>
              </div>

              <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
                <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
                  <a href="#" className={`${styles.button}`} onClick={() => this.EditItem()}>
                    <span className={styles.label}>Update item</span>
                  </a>&nbsp;
                  <a href="#" className={`${styles.button}`} onClick={() => this.DeleteItem()}>
                    <span className={styles.label}>Delete item</span>
                  </a>
                </div>
              </div>

              <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
                <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
                  {/* Display the status of the state*/}
                  {this.state.status}
                  <ul>
                    {/* Call Item JSX element to render list item details */}
                    {items}
                  </ul>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // Function is used to Add new item in SharePoint list.
  private AddItem() {
    alert("Add Button clicked!!");
    if(Environment.type === EnvironmentType.SharePoint){
      this.setState({
        status:"Creating new item.....",
        items:[]
      });

      const body:string = JSON.stringify({
        'Title' : `Item ${new Date()}`
      });

      // Add Post call of rest API to Add item in List

      this.props.spHttpClient.post(
      `${this.props.siteURL}/_api/web/lists/getbytitle('${this.props.ListName}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type':'application/json;odata=nometadata',
          'odata-version': ''
        },
        body : body
      }).then((response:SPHttpClientResponse):Promise<IListItem>=>{
        return response.json();
      }).then((item:IListItem):void=>{
        this.setState({
          status:`Item '${item.Title}' (ID: ${item.ID}) successfully Added!!`,
          items:[]
        });
        },(error:any):void=>{
          this.setState({
            status:'Error while adding new item ' + error,
            items:[]
          });
      });
    }
    else
    {
       this.setState({
         status:"Please connect to SharePoint Online enviornment.You are running in local server",
         items:[]
       });
    }

 }

  private GetItems(){
    alert("Read Button clicked!!");

    this.setState({
      status:'Loading items...',
      items:[]
    });
    if(Environment.type === EnvironmentType.SharePoint){
    this.getLatestItemId().then((itemId:number):Promise<SPHttpClientResponse>=>{
      if(itemId===-1){
         throw new Error('No Items found in list');
      }
      this.setState({
        status:`Loading information abount item ID: ${itemId}...`,
        items:[]
      });

      return this.props.spHttpClient.get(
        `${this.props.siteURL}/_api/web/lists/
        getbytitle('${this.props.ListName}')/items(${itemId})?
        $select=Title,Id,`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
    }).then((res:SPHttpClientResponse):Promise<IListItem>=>{
      return res.json();
    }).then((item:IListItem):void=>{
      this.setState({
        status:`Item ID: ${item.ID}, Title : ${item.Title}`,
        items:[]
      });
    },(error:any):void=>{
      this.setState({
        status:`Loading latest item with error : ` + error,
        items:[]
      });
    });
  }
  else{
    this.setState({
      status:"Please connect to SharePoint Online enviornment.You are running in local server",
      items:[]
    });
  }
  }

  private EditItem() {
    alert("Edit Button clicked!!");

    if(Environment.type === EnvironmentType.SharePoint){
      this.setState({
        status: 'Loading latest items...',
        items: []
      });

      let latestItemId: number = undefined;

      this.getLatestItemId()
        .then((itemId: number): Promise<SPHttpClientResponse> => {
          if (itemId === -1) {
            throw new Error('No items found in the list');
          }

          latestItemId = itemId;
          this.setState({
            status: `Loading information about item ID: ${latestItemId}...`,
            items: []
          });

          return this.props.spHttpClient.get(`${this.props.siteURL}/_api/web/lists/getbytitle('${this.props.ListName}')/items(${latestItemId})?$select=Title,Id`,
           SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'odata-version': ''
              }
            });
        })
        .then((response: SPHttpClientResponse): Promise<IListItem> => {
          return response.json();
        })
        .then((item: IListItem): void => {
          this.setState({
            status: 'Loading latest items...',
            items: []
          });

          const body: string = JSON.stringify({
            'Title': `Updated Item ${new Date()}`
          });

          this.props.spHttpClient.post(`${this.props.siteURL}/_api/web/lists/getbytitle('${this.props.ListName}')/items(${item.ID})`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata',
                'odata-version': '',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE'
              },
              body: body
            })
            .then((response: SPHttpClientResponse): void => {
              this.setState({
                status: `Item with ID: ${latestItemId} successfully updated`,
                items: []
              });
            }, (error: any): void => {
              this.setState({
                status: `Error updating item: ${error}`,
                items: []
              });
            });
        });
    }
    else
    {
      this.setState({
        status:"Please connect to SharePoint Online enviornment.You are running in local server",
        items:[]
      });
    }

  }
  private DeleteItem() {
    alert("Delete Button clicked!!");
    if(Environment.type === EnvironmentType.SharePoint){
      if(!window.confirm("Are you sure want to delete this item?")){
        return;
      }
      else
      {
        this.setState({
          status: 'Loading latest items...',
          items: []
        });

        let latestItemId:number = undefined;
        let etag:string = undefined;


        this.getLatestItemId().then((itemId:number):Promise<SPHttpClientResponse>=>{
          if(itemId === -1){
            throw new Error("'No items found in the list");
          }

          latestItemId = itemId;

          this.setState({
            status:`Loading information about item id : ${latestItemId}`,
            items:[]
          });

          return this.props.spHttpClient.get(
            `${this.props.siteURL}/_api/web/lists/getbytitle('${this.props.ListName}')/items
            ('${latestItemId}')?$select=Id,Title`,
            SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
        }).then((response:SPHttpClientResponse):Promise<IListItem>=>{
          etag = response.headers.get('ETag');
          return response.json();
        }).then((item:IListItem):Promise<SPHttpClientResponse>=>{
          this.setState({
            status:`Deleting item with id : ${latestItemId}`,
            items:[]
           });

           return this.props.spHttpClient.post(
             `${this.props.siteURL}/_api/web/lists/getbytitle('${this.props.ListName}')/items(${item.ID})`,
             SPHttpClient.configurations.v1,
             {
               headers:{
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata',
                'odata-version': '',
                'IF-MATCH': etag,
                'X-HTTP-Method': 'DELETE'
               }
             });
        }).then((response:SPHttpClientResponse):void=>{
          this.setState({
            status: `Item with ID: ${latestItemId} successfully deleted`,
            items: []
          });
        },(error:any):void=>{
          this.setState({
            status: `Error Deleting item: ${error}`,
            items: []
          });
        });
      }
    }
    else
    {
      this.setState({
        status:"Please connect to SharePoint Online enviornment.You are running in local server",
        items:[]
      });
    }
  }

  // get latest Item from the List
  // Function is used to call rest api to get the latest 1 item from the list which is set in property pane.
  private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      this.props.spHttpClient.get(`${this.props.siteURL}/_api/web/lists/getbytitle('${this.props.ListName}')/items?$orderby=Id desc&$top=1&$select=id`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { value: { Id: number }[] }): void => {
          if (response.value.length === 0) {
            resolve(-1);
          }
          else {
            resolve(response.value[0].Id);
          }
        });
    });
  }
}
