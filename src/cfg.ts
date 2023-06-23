import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Lists.Announcements,
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            CustomFields: [
                {
                    name: "BackgroundColor",
                    title: "Background Color",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "Content",
                    title: "Content",
                    type: Helper.SPCfgFieldType.Note,
                    required: true,
                    noteType: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "ExpirationDate",
                    title: "Expiration Date",
                    type: Helper.SPCfgFieldType.Date,
                    required: true,
                    displayFormat: SPTypes.DateFormat.DateOnly
                } as Helper.IFieldInfoDate,
                {
                    name: "LinkUrl",
                    title: "Link Url",
                    type: Helper.SPCfgFieldType.Url
                },
                {
                    name: "TextColor",
                    title: "Text Color",
                    type: Helper.SPCfgFieldType.Text
                }
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "Content", "ExpirationDate", "BackgroundColor", "TextColor", "LinkUrl"
                    ]
                },
                {
                    ViewName: "Active Items",
                    ViewQuery: '<Where><Geq><FieldRef Name="ExpirationDate" /><Value Type="DateTime"><Today /></Value></Geq></Where>',
                    ViewFields: [
                        "LinkTitle", "Content", "ExpirationDate", "BackgroundColor", "TextColor", "LinkUrl"
                    ]
                }
            ]
        }
    ]
});