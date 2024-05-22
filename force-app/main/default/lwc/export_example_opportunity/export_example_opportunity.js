import { LightningElement, track, api } from 'lwc';
import methodGetOpportunity from '@salesforce/apex/ExportExampleController.getOpportunity';
import { exportXLSX } from 'c/export_xlsx';

export default class Export_example_opportunity extends LightningElement{

    @track opportunity;

    // get some mock data
    connectedCallback(){
        methodGetOpportunity()
        .then(result =>{
            this.opportunity = result;
        });
    }

    // single export
    handleExport(){
        let config = new Array();

        config.push(this.getXLSConfig());

        exportXLSX(config, this.opportunity.Name);
    }

    // to expose for a parent export
    @api
    getXLSConfig(){
        return {
            table: this.refs.opportunityTable,
            tabName: 'Opportunity',
            displayGrid: true,
            zoom: 120,
            style: this.getXLSStyle()
        };
    }

    // style for xml file
    getXLSStyle(){
        return `
            <Style ss:ID="o-default">
                <Alignment ss:Vertical="Center" ss:Horizontal="Left"/>
                <Font ss:Size="12" ss:FontName="Arial"/>
                <Borders>
                    <Border ss:Color="#000000" ss:Weight="1" ss:LineStyle="Continuous" ss:Position="Top"/>
                    <Border ss:Color="#000000" ss:Weight="1" ss:LineStyle="Continuous" ss:Position="Right"/>
                    <Border ss:Color="#000000" ss:Weight="1" ss:LineStyle="Continuous" ss:Position="Bottom"/>
                    <Border ss:Color="#000000" ss:Weight="1" ss:LineStyle="Continuous" ss:Position="Left"/>
                </Borders>
            </Style>
            <Style ss:ID="o-col" ss:Parent="o-default">
                <Font ss:Size="12" ss:FontName="Arial" ss:Bold="1"/>
                <Interior ss:Color="#bfbf67" ss:Pattern="Solid"/>
            </Style>
            <Style ss:ID="o-value" ss:Parent="o-default">
                <Interior ss:Color="#e0e0e0" ss:Pattern="Solid"/>
                <Alignment ss:Vertical="Center" ss:Horizontal="Left"/>
            </Style>
            <Style ss:ID="o-value-money" ss:Parent="o-default">
                <Alignment ss:Vertical="Center" ss:Horizontal="Right"/>
                <NumberFormat ss:Format="Currency"></NumberFormat>
                <Interior ss:Color="#e0e0e0" ss:Pattern="Solid"/>
            </Style>
        `;
    }

}