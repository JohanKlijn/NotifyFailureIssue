import { useEffect, useState } from "react";
import { app } from "@microsoft/teams-js";

export default function Tab() {
  const [error, setError] = useState<string>();
  const [loading, setLoading] = useState<boolean>(true);

  useEffect(() => {

    const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

    const initializeApp = async () => {
      app.initialize()
      .then(() => {
        setLoading(false)
      })
      .catch((reason) =>{
        setError(reason)
      })
    }    
    
    wait(3 * 1000).then(() => {
      initializeApp()
      .then(() => {
        console.log({"isInitialized" : app.isInitialized()})
        app.notifyFailure({reason: app.FailedReason.Other, message: "My custom failure message."}); 
      })     
    })

  }, []);

  return (
    <div >
      My Teams Tab App
    </div>
  );
}
