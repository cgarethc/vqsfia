/* jshint esversion: 6 */

import React from 'react';
import ReactDOM from 'react-dom';
import AppBar from 'material-ui/AppBar';
import Paper from 'material-ui/Paper';
import request from 'superagent';
import {deepOrange500} from 'material-ui/styles/colors';
import getMuiTheme from 'material-ui/styles/getMuiTheme';
import MuiThemeProvider from 'material-ui/styles/MuiThemeProvider';
import SelectField from 'material-ui/SelectField';
import MenuItem from 'material-ui/MenuItem';
import injectTapEventPlugin from 'react-tap-event-plugin';
import Slider from 'material-ui/Slider';
import RaisedButton from 'material-ui/RaisedButton';

injectTapEventPlugin();

const roles = ['Software Engineer', 'Test Engineer', 'Support Engineer', 'User Experience Designer'];
const levels = ['Graduate', 'Junior', 'Intermediate', 'Senior', 'Principal'];

const muiTheme = getMuiTheme({
  palette: {
    accent1Color: deepOrange500,
  },
});

const style = {
  margin: 12,
};

const RoleBox = React.createClass({
  getInitialState(){
    return {value:this.props.initialSelection};
  },
  handleChange: function(event, index, value){
     this.setState({value});
     this.props.onChange(this.props.field, roles[value]);
  },
  render: function(){
    return (
      <SelectField value={this.state.value} onChange={this.handleChange}>
        <MenuItem value={0} primaryText={roles[0]} />
        <MenuItem value={1} primaryText={roles[1]} />
        <MenuItem value={2} primaryText={roles[2]} />
        <MenuItem value={3} primaryText={roles[3]} />
      </SelectField>
    );
  }
});

const LevelBox = React.createClass({
  getInitialState(){
    return {value:this.props.initialSelection};
  },
  handleChange: function(event, index, value){
     this.setState({value});
     this.props.onChange(this.props.field, levels[value]);
  },render: function(){
    return(
      <SelectField value={this.state.value} onChange={this.handleChange}>
        <MenuItem value={0} primaryText={levels[0]} />
        <MenuItem value={1} primaryText={levels[1]} />
        <MenuItem value={2} primaryText={levels[2]} />
        <MenuItem value={3} primaryText={levels[3]} />
        <MenuItem value={4} primaryText={levels[4]} />
      </SelectField>
    );
  }
});

const VqSFIA = React.createClass({
  handleProfile: function(){
    window.location.href=`/form?fromrole=${state.fromrole}&levelfrom=${state.levelfrom}`;
  },
  handleMatrix: function(){
    window.location.href=`/form?fromrole=${state.fromrole}&levelfrom=${state.levelfrom}&torole=${state.torole}&progressionlevel=${state.progressionlevel}`;
  },
  render: function(){
    return (
      <div>
        <AppBar
        title="Progression forms"
        iconElementLeft={<span/>}
        />
        <h3>Current role</h3>
        <div>
          <RoleBox field="fromrole" onChange={changed} initialSelection={0}/>
          <LevelBox field="levelfrom" onChange={changed} initialSelection={0}/>
          <RaisedButton label="Profile" primary={true} style={style} onTouchTap={this.handleProfile}/>
          <h3>Next role</h3>
          <RoleBox field="torole" onChange={changed} initialSelection={0}/>
          <LevelBox field="progressionlevel" onChange={changed} initialSelection={1} />
          <RaisedButton label="Matrix" primary={true} style={style} onTouchTap={this.handleMatrix}/>
        </div>
      </div>
    );
  }
});

let state = {
  fromrole: 'Software Engineer',
  levelfrom: 'Graduate',
  torole: 'Software Engineer',
  progressionlevel: 'Junior'
};

function changed(field, value){
  console.log(field + ': ' + value);
  state[field] = value;
}

ReactDOM.render(
  <MuiThemeProvider muiTheme={getMuiTheme()}><VqSFIA/></MuiThemeProvider>,
  document.getElementById('content')
);
