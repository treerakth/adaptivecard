// Import Modules
const axios = require('axios');
const fs = require('fs'); // Add this line to import the 'fs' module

// Webhook URL ของ Microsoft Teams
const WEBHOOK_URL = "https://nu365.webhook.office.com/webhookb2/efb3b6cb-58c9-4b21-a4e3-c34adfa86609@bf1eb3f8-19d2-409d-b0c5-4e80c943fd52/IncomingWebhook/9bf57115d7a3478d83155dfb091b57f7/395194ba-631a-4986-8e13-827d36d58e85/V2p9-r_PPqfG3_boA8VcG_PadiaucvOh73hF5LvKVRVCQ1";

// ฟังก์ชันส่งข้อความไปยัง Microsoft Teams
async function sendMessageToTeams() {
    try {
        // Payload รูปแบบ Adaptive Card
        const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));
        const color = JSON.parse(fs.readFileSync('color.json', 'utf8'));
        const requestDetails = data.requestDetails || []; // Ensure requestDetails is defined
        const cardContent = {
            type: "message",
            attachments: [
                {
                    contentType: "application/vnd.microsoft.card.adaptive",
                    content: {
                        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                        type: "AdaptiveCard",
                        version: "1.4",
                        body: [
                            {
                                type: "Image",
                                url: data.imageUrl,
                                size: "Large",
                                style: "Default",
                                altText: "Header Image",
                                height: "170px",
                                horizontalAlignment: "Center"
                            },
                            {
                                type: "TextBlock",
                                text: "เลขที่ใบคำขอ",
                                size: "Medium",
                                weight: "Bolder",
                                horizontalAlignment: "Center",
                                spacing: "Medium"
                            },
                            {
                                type: "TextBlock",
                                text: data.requestNumber,
                                size: "ExtraLarge",
                                weight: "Bolder",
                                color: "Default",
                                horizontalAlignment: "Center",
                                spacing: "Small"
                            },
                            {
                                type: "TextBlock",
                                text: "(คำขอใช้รถนอกเงื่อนไข)",
                                horizontalAlignment: "Center",
                                color: color.red,
                                weight: "Bolder",
                                spacing: "Small"
                            },
                            {
                                type: "Container",
                                style: "emphasis",
                                items: [
                                    {
                                        type: "TextBlock",
                                        text: "โปรดอนุมัติใบคำขอภายในวันที่",
                                        horizontalAlignment: "Center",
                                        color: color.red,
                                        weight: "Bolder",
                                        spacing: "Medium"
                                    },
                                    {
                                        type: "TextBlock",
                                        text: data.approvalDeadline,
                                        horizontalAlignment: "Center",
                                        color: color.green,
                                        weight: "Bolder",
                                        spacing: "Small"
                                    },
                                    {
                                        type: "TextBlock",
                                        text: "ผู้อนุมัติ คุณ " + data.approver,
                                        horizontalAlignment: "Center",
                                        weight: "Bolder",
                                        spacing: "Medium"
                                    }
                                ]
                            },
                            {
                                type: "Container",
                                items: [
                                    {
                                        type: "TextBlock",
                                        text: "ข้อมูลใบคำขอ",
                                        size: "Medium",
                                        weight: "Bolder",
                                        spacing: "Medium"
                                    },
                                    {
                                        type: "TextBlock",
                                        text: "1. วันที่เดินทาง:" + data.DateTravel,
                                        wrap: true,
                                        spacing: "Small"
                                    },
                                    {
                                        type: "TextBlock",
                                        text: "2. สถานที่ปลายทาง: " + data.TravelProvice,
                                        wrap: true,
                                        spacing: "Small"
                                    },
                                    {
                                        type: "TextBlock",
                                        text: "3. สถานที่นัดหมาย: " + data.Appointment,
                                        wrap: true,
                                        spacing: "Small"
                                    },
                                    {
                                        type: "TextBlock",
                                        text: "4. ประเภทรถ:",
                                        wrap: true,
                                        spacing: "Small"
                                    },
                                    ...requestDetails.map(detail => ({
                                        type: "TextBlock",
                                        text: detail,
                                        wrap: true,
                                        spacing: "Small"
                                    }))
                                ]
                            }
                        ],
                        actions: [
                            {
                                type: "Action.Submit",
                                title: "ดูใบงาน",
                                data: { action: "view" }
                            },
                            {
                                type: "Action.Submit",
                                title: "แก้ไขใบคำขอ",
                                data: { action: "edit" }
                            },
                            {
                                type: "Action.Submit",
                                title: "ยกเลิกใบคำขอ",
                                data: { action: "cancel" }
                            }
                        ]
                    }
                }
            ]
        };

        // ส่งข้อมูลผ่าน HTTP POST
        const response = await axios.post(WEBHOOK_URL, cardContent, {
            headers: {
                'Content-Type': 'application/json'
            }
        });

        if (response.status === 200) {
            console.log("Message sent successfully to Teams!");
        } else {
            console.error(`Failed to send message: ${response.status}, ${response.data}`);
        }
    } catch (error) {
        console.error("Error sending message to Teams:", error);
    }
}

// เรียกใช้งานฟังก์ชัน
sendMessageToTeams();
