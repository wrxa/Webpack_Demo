import React, {Component} from 'react';
import config from './config.json';

// 声明组件
class ShowSXBackgroud extends Component {
    constructor(props){
        super(props);
        // 声明并设定变量的state
        this.state={
            'buttonName':config.showSXBackgroudButtonText,
            'buttonClass':'btn btn-primary',
            'lifemarksx_img_src':""
        };
    }

    // 按钮事件
    changeButtonText() {
        this.setState({buttonName: config.showSXBackgroudButtonText_afterClick}); 
        this.setState({buttonClass: "btn-success"}); 
        this.setState({lifemarksx_img_src: config.img_url}); 
    }

    // 渲染组件
    render() {
        var buttonName = this.state.buttonName;
        var buttonClass = this.state.buttonClass;
        var lifemarksx_img_src = this.state.lifemarksx_img_src;
        return (
            <div>
                <div class="container">
                    {/* <button onClick={this.changeButtonText.bind(this)} type="button" class="btn btn-primary">{config.showSXBackgroudButtonText}</button> */}
                    <button onClick={this.changeButtonText.bind(this)} type="button" class={buttonClass}>{buttonName}</button>
                </div>
                <br />
                <div>
                    <img src={lifemarksx_img_src} class="img-fluid"></img>
                </div>
            </div>
        );
    }
}

// 对外开放class
export default ShowSXBackgroud
