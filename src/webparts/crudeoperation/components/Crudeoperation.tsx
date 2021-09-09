import * as React from 'react';
import * as ReactDOM from 'react-dom';
import Modal from 'react-responsive-modal';
import { ICrudeoperationProps } from './ICrudeoperationProps';
//import 'react-responsive-modal/lib/react-responsive-modal.css';
import 'react-responsive-modal/styles.css';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import styles from './Crudeoperation.module.scss';
import { Web, IWeb, sp } from "@pnp/sp/presets/all";
import "./styles/style.css";

const share: any = require('./styles/share.png');

export default class Crudeoperation extends React.Component<ICrudeoperationProps, any> {

  constructor(props){
    super(props)
    this.state = {
      open: false,
      fileShare:false,
      itemID:0,
      itemColl:[]
    };
    this.getFileItems=this.getFileItems.bind(this)
  }

  componentDidMount(){
    this.getFileItems();
    
  }
  onOpenModal = (itemID) => {
    //alert('Opened')
    this.setState({ open: true });
    this.setState({ itemID: itemID });
  };

  onCloseModal = () => {
    this.setState({ open: false });
  };

  private _getToPeoplePickerItems(items: any[]) {
    debugger;
    console.log('Items:', items);
    let userarr = [];
     let idarr = [];
    let defaultSelectedUsers = [];
    let usernamearr: string[] = [];
    items.forEach(user => {
      userarr.push({ ID: user.id, LoginName: user.loginName });
     // defaultSelectedUsers.push(user.loginName);
     idarr.push(user.id);
     usernamearr.push( user.loginName.split('|membership|')[1].toString() );
    });
    this.setState({ toSelectedUsers : userarr});
    
  }

  public onSave(){
    debugger;
    this.setState({fileShare:true});
    let usernamearr : any[] =[];
    let items: any[] = [];
    let namearr;
    
    let idarr : any[] = [];
    items = this.state.toSelectedUsers;
    items.forEach(user => {
     // userarr.push({ ID: user.id, LoginName: user.loginName });
     // defaultSelectedUsers.push(user.loginName);
     usernamearr.push(user.LoginName.split('|membership|')[1].toString());
    idarr.push(user.ID);
    });
    namearr = usernamearr.join(';').toString();
    this.setState({toUsers:namearr,open:false});
    
    //Update List Item
    sp.web.lists.getByTitle('MyDoc').items.getById(parseInt(this.state.itemID)).update({
      "SharedWith": this.state.toUsers
    }).then(()=>{
      console.log('Updated List');
      alert('Successfully shared');
    })
    setTimeout(() => {
      this.setState({fileShare:false});
    }, 5000);
  }

  private getFileItems(){
    let { itemColl } = this.state;
    sp.web.lists.getByTitle('MyDoc').items.select('Id,FileRef,File,Title').expand('File').get().then(file=>{
      debugger;
      console.log(file);
      file.map((item)=>{       
        console.log(item.FileRef);
        itemColl.push({FileName:item.File.Name,FileURL:item.FileRef,ItemID:item.ID});       
      }) 
      //file[0].File.Name
      //file[0].FileRef
      this.setState({ open: false });
    })
  }

  public render(): React.ReactElement<ICrudeoperationProps> {
    const { open, itemColl,fileShare } = this.state;

    return (
      <div className="form-container">
      {fileShare ? <div className="fileSuccess"><ul><li>File has been shared successfully.</li></ul></div> :''}
      <div className="form-heading">My Documents</div>
      {itemColl.map((item)=>{
      debugger;
      return(
        <div className="txt-lbl-card">
        <div className="form-row">
          <div className="form-col50">
            <div className="form-row">
              <div className="form-label"><a href={item.FileURL}>{item.FileName}</a></div>
              <div className="form-label-value">
              <img className="image" onClick={() => this.onOpenModal(item.ItemID)} src={share}></img>
              </div>
            </div>
          </div>
        </div>
      </div>)
    })}

      {/* <div className="txt-lbl-card">
        <div className="form-row">
          <div className="form-col50">
            <div className="form-row">
              <div className="form-label">File Name:</div>
              <div className="form-label-value">
              <img className="image" onClick={() => this.onOpenModal('sdf')} src={share}></img>
              </div>
            </div>
          </div>
        </div>
      </div> */}

    {/*Modal Component Script */}
    <Modal open={open} onClose={this.onCloseModal} center>
          <h2>Please select users to Share the document</h2>
          <PeoplePicker    
                context={this.props.context}    
                titleText="Select User"    
                personSelectionLimit={25}    
                groupName={""} // Leave this blank in case you want to filter from all users    
                //showtooltip={true}    
                suggestionsLimit={20}
                disabled={false}   
                required={true} 
                ensureUser={true}   
                onChange={(items)=>this._getToPeoplePickerItems(items)}    
                //showHiddenInUI={false}    
                principalTypes={[PrincipalType.User]}    
                resolveDelay={500} /> 
                <div className="buttonPosition">
                  <button className="confirmStyle" type="button" onClick={(e)=>this.onSave()}>Confirm</button>
                </div>
        </Modal>



    </div>
    );
  }
}
