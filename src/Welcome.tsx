import { AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import { Button, Container, Modal } from 'react-bootstrap';
import 'bootstrap/dist/css/bootstrap.min.css';
import './Welcome.css';
import { RouteComponentProps } from 'react-router-dom';
import { AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { useAppContext } from './AppContext';
import { getMail } from './GraphService';

import { useEffect, useState } from 'react';
import $ from 'jquery';

  export default function Welcome(props: RouteComponentProps) {
    let targetParent: any = null;
    const app = useAppContext();
    console.log(app);
    const [mails, setMails] = useState([]);

    const [modalState, setModalState] = useState(false);
    const [modalHeader, setModalHeader] = useState("");
    const [modalBody, setModalBody] = useState("");

    let showModal = () => {
      setModalState(true);
    }

    let hideModal = () => {
      console.log("Hiding modal")
      setModalState(false);
    }

    useEffect(() => {
      ( async () => {
        try {
          if(app.authProvider){
            const mails = await getMail(app.authProvider)
            setMails(mails.value)
            $("h4").on("click", function(){
              $("p").css("color", "blue");
            });
          }
          
          
        } catch (error) {
          console.log(error)
        }
      })()
      
    }, []);

    let emailClicked = (p: any) => {
      showModal();
      setModalHeader(p.target.innerText);
      setModalBody(p.target.nextSibling.innerText);
      p.target.closest(".mail").style.display = "None";
    }

    return (
      <div className="p-5 mb-4 bg-light rounded-3">
        <Container fluid>
          <AuthenticatedTemplate>
            <div>
              <h4>Welcome {app.user?.displayName || ''}!</h4>
              <p>Click any email below.</p>

              <Modal show={modalState} >
                <Modal.Header><h3 className="mail-modal-header">{modalHeader}</h3> <button onClick={hideModal} className="btn btn-danger">X</button></Modal.Header>
                <Modal.Body><div className="mail-modal-body">{modalBody}</div></Modal.Body>
                </Modal>

              <div className="mail-list">
                  <h1>Your Mail</h1><hr />
                  <div className="mail-list">
                    {mails.map((mail : any, index) => {
                      return (
                        <div key={index} className="mail">
                          <h2 className="mail-subject"  onClick={(e) => emailClicked(e)}>{mail.subject}</h2>
                          <div className='mail-message' style={{display: "none"}}><p>{mail.bodyPreview}</p></div>
                        </div>
                      )
                    })}
                    
                  </div>
              </div>
            </div>
          </AuthenticatedTemplate>
          <UnauthenticatedTemplate>
            <h1>SignIn to view your Mail!!!</h1>
            <Button color="primary" onClick={app.signIn!}>Click here to sign in</Button>
          </UnauthenticatedTemplate>
        </Container>
      </div>
    );
  }