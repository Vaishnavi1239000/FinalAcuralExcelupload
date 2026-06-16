import * as React from 'react';
import './Site.scss'

import ManageAccess from './ManageAccess'
import UploadAccrual from './UploadAccrual'
import { sp } from "../DataUploadWebPart";

import { IDataUploadProps } from './IDataUploadProps';

export default function AccrualSheet(props: IDataUploadProps) {

  const [page, setPage] = React.useState("home");
  const [hasAccess, setHasAccess] = React.useState<boolean | null>(null);


  const checkUserAccess = async () => {

    try {

      const user = await sp.web.currentUser();

      const today = new Date();

      const items: any = await sp.web.lists
        .getByTitle("AccrualSheetAccessList")
        .items
        .select("FromDate", "UptoDate", "Username/EMail")
        .expand("Username")
        .top(5000)();

      const access = items.some((item: any) => {

        const from = new Date(item.FromDate);
        const upto = new Date(item.UptoDate);

        return (
          item.Username?.EMail === user.Email &&
          today >= from &&
          today <= upto
        );

      });

      setHasAccess(access);

    } catch (error) {

      console.log(error);
      setHasAccess(false);

    }

  };
  React.useEffect(() => {
    void checkUserAccess();
  }, []);

  // Loading
  if (hasAccess === null) {
    return <h3>Loading...</h3>;
  }

  // No Access
  if (!hasAccess) {
    return (
      <div style={{ padding: "40px" }}>
        <h2 style={{ color: "red" }}>Access Denied</h2>
        <p>You do not have access to this portal.</p>
      </div>
    );
  }

  if (page === "manage") {
    return <ManageAccess {...props} />;
  }

  if (page === "upload") {
    return <UploadAccrual />;
  }

  return (
    <div className="container-fluid">

      <div className='header'>
        <h1>Accrual Sheet</h1>
      </div>
      <div className="PaddAll">
        <div className="buttonContainer">

          <button
            className="btn"
            onClick={() => setPage("manage")}
          >
            Manage Access
          </button>

          <button
            className="btn"
            onClick={() => setPage("upload")}
          >
            Upload Accrual Sheet
          </button>

          <button
            className="btn"
            onClick={() => setPage("report")}
          >
            Adjustment Report
          </button>

        </div>
      </div>
    </div>
  );

}