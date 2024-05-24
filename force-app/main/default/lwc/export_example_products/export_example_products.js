import { LightningElement, track, api } from 'lwc';
import methodGetProducts from '@salesforce/apex/ExportExampleController.getProducts';
import { exportXLSX } from 'c/export_xlsx';

export default class Export_example_products extends LightningElement{

    @track products;

    // get some mock data
    connectedCallback(){
        methodGetProducts()
        .then(result =>{
            this.products = result;

            let index = 0;

            this.products.forEach(product =>{
                product.key = index;

                index ++;
            });
        });
    }

    // single export
    handleExport(){
        let config = new Array();

        config.push(this.getXLSConfig());

        exportXLSX(config, 'Table Products');
    }

    // to expose for a parent export
    @api
    getXLSConfig(){
        return {
            table: this.refs.productsTable,
            tabName: 'Products',
            displayGrid: false,
            zoom: 90,
            style: this.getXLSStyle(),
            columns:[
                {width: 150},
                {width: 100},
                {width: 50},
                {width: 50, hidden: true},
            ]
        };
    }

    // style for xml file
    getXLSStyle(){
        return `
            <Style ss:ID="p-default">
                <Alignment ss:Vertical="Center" ss:Horizontal="Center"/>
                <Font ss:Size="12" ss:FontName="Calibri"/>
                <Borders>
                    <Border ss:Color="#000000" ss:Weight="1" ss:LineStyle="Continuous" ss:Position="Top"/>
                    <Border ss:Color="#000000" ss:Weight="1" ss:LineStyle="Continuous" ss:Position="Right"/>
                    <Border ss:Color="#000000" ss:Weight="1" ss:LineStyle="Continuous" ss:Position="Bottom"/>
                    <Border ss:Color="#000000" ss:Weight="1" ss:LineStyle="Continuous" ss:Position="Left"/>
                </Borders>
            </Style>
            <Style ss:ID="p-table-title">
                <Alignment ss:Vertical="Center" ss:Horizontal="Left"/>
                <Font ss:Size="20" ss:FontName="Calibri" ss:Bold="1"/>
            </Style>
            <Style ss:ID="p-col-title" ss:Parent="p-default">
                <Font ss:Size="12" ss:FontName="Calibri" ss:Bold="1"/>
                <Interior ss:Color="#C0C0C0" ss:Pattern="Solid"/>
            </Style>
            <Style ss:ID="p-col-value-text" ss:Parent="p-default">
                <Alignment ss:Vertical="Center" ss:Horizontal="Left"/>
            </Style>
            <Style ss:ID="p-col-value-money" ss:Parent="p-default">
                <Alignment ss:Vertical="Center" ss:Horizontal="Right"/>
                <NumberFormat ss:Format="Currency"></NumberFormat>
            </Style>
            <Style ss:ID="p-col-value-percent" ss:Parent="p-default">
                <Alignment ss:Vertical="Center" ss:Horizontal="Right"/>
                <NumberFormat ss:Format="Percent"></NumberFormat>
            </Style>
        `;
    }

}