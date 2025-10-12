// client.js (Revised, based on previous recommendation)

import React, { useEffect, useRef } from "react";

export default function WordEditor() {
    const fileId = "sample.docx";
    const token = "e497b0b359d0b35065d022a5f71df00c03d0291d307f64d4866fd8aa0243b26f";

    // Target URL for the Office Editor (must match the discovery action URL)
    const POST_ACTION_URL = "https://word-edit.officeapps.live.com/we/wordeditorframe.aspx";

    // The WOPISrc URL that your WOPI host (ngrok) is listening on
    // NOTE: Replace a291020c90a5.ngrok-free.app with your current ngrok URL!
    const wopiSrc = `https://breezy-glasses-fetch.loca.lt/wopi/files/${fileId}?access_token=${token}`;

    const formRef = useRef(null);

    useEffect(() => {
        // Automatically submit the form once the component mounts
        if (formRef.current) {
            console.log("Submitting WOPI POST form...");
            formRef.current.submit();
        }
    }, []);

    return (
        <div>
            {/* 1. Define the Iframe (target for the POST) */}
            <iframe
                title="Word Online"
                id="office-iframe"
                name="office-iframe" // <-- NAME must match FORM TARGET
                width="100%"
                height="800px"
                frameBorder="0"
            />

            {/* 2. Define the hidden POST Form */}
            <form
                ref={formRef}
                action={POST_ACTION_URL}
                method="post"
                target="office-iframe" // <-- TARGET must match IFRAME NAME
                style={{ display: 'none' }}
            >
                {/* 3. Pass WOPISrc as a hidden input field */}
                <input type="hidden" name="WOPISrc" value={wopiSrc} />
            </form>
        </div>
    );
}