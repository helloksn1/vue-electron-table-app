const fs = require("fs");
const { ipcRenderer } = require("electron");
const xlsx = require('node-xlsx').default;
const { isNumber } = require("node-xlsx/lib/helpers");

const getEnd = (start, dur) => {
    const end = moment(start);
    if (dur === 'week') end.add(7, 'days');
    else if (dur === 'month') end.add(1, 'months');
    else if (dur === 'season') end.add(3, 'months');
    else if (dur === 'half') end.add(6, 'months');
    else if (dur === 'year') end.add(1, 'years');
    else end.add(2, 'years');
    return end.format("YYYY-MM-DD");
}
const getBkColor = (left, dur, state) => {
    let result = '';
    if ((left <= 1 && dur == 'week') || (left <= 3 && dur != 'week'))
        result = 'left-error';
    else if ((left <= 3 && dur == 'week') || (left <= 10 && dur != 'week'))
        result = 'left-warning';
    if (result === '') return result;
    if (state === 'user')
        result = 'left-user-checked';
    else if (state === 'admin')
        result = 'left-admin-checked';
    return result;
}

const mainCont = {
    data() {
        let dat;
        if (fs.existsSync("dat")) {
            dat = JSON.parse(fs.readFileSync("dat"));
            for (let t of dat.data)
                t.editing = false;
        }
        else
            dat = {
                admin_password: 'admin',
                normal_password: 'pass',
                saved_admin_password: '',
                saved_normal_password: '',
                usertype: 'normal',
                data: [{
                    name: "空表",
                    data: [{}]
                }]
            };

        return {
            logedIn: false,
            admin_password: dat.admin_password,
            normal_password: dat.normal_password,
            saved_admin_password: dat.saved_admin_password,
            saved_normal_password: dat.saved_normal_password,
            cur_table: 0,
            deleting: -1,
            data: dat.data,
            wrongpass: false,
            wdataShow: false,
            settingsShow: false,
            sound: '',
            cleanfilter: '',
            repairfilter: '',
            usertype: 'normal'
        };
    },
    computed: {
        tableNames() {
            const res = [];
            for (let i = 0; i < this.data.length; i++) {
                const table = this.data[i];
                res.push({
                    name: table.name,
                    selected: i == this.cur_table ? "selected" : "",
                    editing: table.editing
                })
            }
            return res;
        },
        cur_table_data() {
            const ctd = [];
            for (let r of this.data[this.cur_table].data) {
                const cr = {};
                for (let k in r) {
                    cr[k] = r[k];
                }
                if (cr.clean_start && cr.clean_duration) {
                    cr.clean_end = getEnd(cr.clean_start, cr.clean_duration);
                    cr.clean_left = moment(cr.clean_end).diff(moment(new Date()), 'days') + 1;
                    cr.clean_left_class = getBkColor(cr.clean_left, cr.clean_duration, cr.clean_check_state);
                }
                if (cr.repair_start && cr.repair_duration) {
                    cr.repair_end = getEnd(cr.repair_start, cr.repair_duration);
                    cr.repair_left = moment(cr.repair_end).diff(moment(new Date()), 'days') + 1;
                    cr.repair_left_class = getBkColor(cr.repair_left, cr.repair_duration, cr.repair_check_state);
                }
                ctd.push(cr);
            }
            return ctd;
        },
        warningData() {
            const today = moment(new Date());
            const wd = [];
            for (let ti = 0; ti < this.data.length; ti++) {
                const tb = this.data[ti];
                for (let ri = 0; ri < tb.data.length; ri++) {
                    const r = tb.data[ri];
                    let clean_end = '', clean_left = 0, clean_col = '';
                    if (r.clean_start && r.clean_duration) {
                        clean_end = getEnd(r.clean_start, r.clean_duration);
                        clean_left = moment(clean_end).diff(today, 'days') + 1;
                        clean_col = getBkColor(clean_left, r.clean_duration, r.clean_check_state);
                    }
                    let repair_end = '', repair_left = 0, repair_col = '';
                    if (r.repair_start && r.repair_duration) {
                        repair_end = getEnd(r.repair_start, r.repair_duration);
                        repair_left = moment(repair_end).diff(today, 'days') + 1;
                        repair_col = getBkColor(repair_left, r.repair_duration, r.repair_check_state);
                    }
                    if (clean_col || repair_col) {
                        wd.push({
                            'ti': ti,
                            'table': tb.name,
                            'ri': ri,
                            'name': r.name,
                            'clean_end': clean_end,
                            'clean_duration': r.clean_duration,
                            'clean_col': clean_col,
                            'repair_end': repair_end,
                            'repair_duration': r.repair_duration,
                            'repair_col': repair_col,
                        })
                    }
                }
            }
            return wd;
        }
    },
    methods: {
        settingsClicked() {
            this.settingsShow = true;
            this.newpass = '';
            this.oldpass = '';
            this.wrongpass = false;
        },
        login(usertype, password, passsaved) {
            if (usertype === 'admin') {
                if (this.admin_password === password) {
                    this.logedIn = true;
                    this.usertype = 'admin';
                    this.saved_admin_password = passsaved ? password : '';
                }
                else
                    this.wrongpass = true;
            }
            else {
                if (this.normal_password === password) {
                    this.logedIn = true;
                    this.usertype = 'normal';
                    this.saved_normal_password = passsaved ? password : '';
                }
                else
                    this.wrongpass = true;
            }
        },
        changed(id, k, v) {
            const td = this.data[this.cur_table].data;
            const rd = td[id];
            rd[k] = v;
            if (k === 'clean_start' || k === 'clean_duration') 
                rd.clean_check_state = '';
            if (k === 'repair_start' || k === 'repair_duration')
                rd.repair_check_state = '';
            if (id === td.length - 1) {
                td.push({});
            }
        },
        recordDelete(id) {
            if (this.data[this.cur_table].data.length > 1)
                this.data[this.cur_table].data.splice(id, 1);
        },
        selectTable(idx) {
            if (this.data[idx].editing) return;
            this.cur_table = idx;
        },
        readFromExcel() {
            ipcRenderer.send('read-path', {});
        },
        addTable() {
            this.data.push({
                name: "新表",
                data: [{}],
                editing: true
            })
        },
        editTable(idx) {
            this.data[idx].editing = true;
            return false;
        },
        deleteTable(idx) {
            this.deleting = idx;
        },
        deleteCheck() {
            if (this.cur_table == this.deleting)
                this.cur_table = 0;
            this.data.splice(this.deleting, 1);
            if (this.data.length == 0) {
                this.data.push({
                    name: "空表",
                    data: [{}]
                })
            }
            this.deleting = -1;
        },
        deleteCancel() {
            this.deleting = -1;
        },
        tableNameChanged(idx, str) {
            this.data[idx].name = str;
        },
        tableNameLocked(idx) {
            this.data[idx].editing = false;
        },
        userCheck(ti, ri, i, type) {
            if (type === 'clean')
                this.data[ti].data[ri].clean_check_state = '';
            if (i == 0)
                this.data[ti].data[ri].clean_check_state = type;
            else
                this.data[ti].data[ri].repair_check_state = type;
        },
        userCheckAll(type) {
            const today = moment(new Date());
            for (let t of this.data) {
                for (let r of t.data) {
                    if (r.clean_start && r.clean_duration) {
                        const clean_end = getEnd(r.clean_start, r.clean_duration);
                        const clean_left = moment(clean_end).diff(today, 'days') + 1;
                        const clean_col = getBkColor(clean_left, r.clean_duration);
                        if (clean_col !== '' && r.clean_check_state !== 'admin')
                            r.clean_check_state = type;
                    }
                    if (r.repair_start && r.repair_duration) {
                        const repair_end = getEnd(r.repair_start, r.repair_duration);
                        const repair_left = moment(repair_end).diff(today, 'days') + 1;
                        const repair_col = getBkColor(repair_left, r.repair_duration);
                        if (repair_col != '' && r.repair_check_state !== 'admin')
                            r.repair_check_state = type;
                    }
                }
            }
        },
        changepass(oldpass, newpass) {
            if (this.usertype === 'admin') {
                if (oldpass !== this.admin_password) {
                    return;
                }
                this.admin_password = newpass;
                this.settingsShow = false;
            }
            else {
                if (oldpass !== this.normal_password) {
                    return;
                }
                this.normal_password = newpass;
                this.settingsShow = false;
            }
        }
    },
    mounted() {
        setInterval(() => {
            const cur = new Date();
            if (cur.getHours() >= 9 && cur.getMinutes() === 0) {
                const dat = this.warningData;
                if (dat.length !== 0) {
                    let error = false, warning = false, userchecked = false;
                    for (let r of dat) {
                        if (r.clean_col === 'left-error' || r.repair_col === 'left-error') {
                            error = true;
                            break;
                        }
                        if (r.clean_col === 'left-warning' || r.repair_col === 'left-warning')
                            warning = true;
                        if (r.clean_col === 'left-user-checked' || r.repair_col === 'left-user-checked')
                            userchecked = true;
                    }
                    if (error || warning || userchecked) {
                        this.wdataShow = true;
                        if (error) this.sound = "alarm0.wav";
                        else if (warning) this.sound = "alarm1.wav";
                        else if (userchecked) this.sound = "alarm2.wav";
                        else this.sound = "";
                    }
                }
            }
            else {
                this.sound = "";
            }
        }, 59000);
    }
};

const loginPad = {
    props: [ 'wrongpass', 'savedadminpass', 'savednormalpass' ],
    emits: [ 'login' ],
    data() {
        return {
            password: this.$props.savedadminpass,
            usertype: 'admin',
            passsaved: this.$props.savedadminpass ? true : false
        }
    },
    methods: {
        userChanged() {
            console.log('changed');
            if (this.usertype == 'admin') {
                this.password = this.$props.savedadminpass;
            }
            else {
                this.password = this.$props.savednormalpass;
            }
            this.passsaved = this.password ? true : false;
        }
    },
    template: `
        <div class="login-pad">
            <div class="black-bk-color"></div>
            <div class="password-pad">
                <div><img class="login-logo" src="logo.png"></div>
                <input type="radio" name="usertype" value="normal" v-model="usertype" @change="userChanged">普通用户
                <input type="radio" name="usertype" value="admin" v-model="usertype" @change="userChanged">管理员
                <input
                    class="password-input"
                    :style="{
                        'border-bottom': wrongpass ? '2px solid red' : '2px solid white'
                    }"
                    type="password"
                    v-model="password"
                    placeholder="请输入密码"
                    @keydown.enter="$emit('login', usertype, password, passsaved)">
                <input type="checkbox" name="passsaved" v-model="passsaved">记住密码
                <button
                    class="password-button"
                    @click="$emit('login', usertype, password, passsaved)"
                >
                    <i class="fa fa-sign-in"></i> 登 录
                </button>
            </div>
        </div>
    `
};

const leftPad = {
    props: [ 'tables', 'cur_table', 'deleting', 'deletecheck', 'deletecancel', 'editable' ],
    emits: [ 'select', 'load', 'edit', 'delete', 'namechanged', 'locked', 'openwdata', 'opensettings' ],
    methods: {},
    template: `
        <div class="left-pad">
            <div class="table-name-container">
                <div
                    class="table-name"
                    v-for="(table, idx) in tables"
                    :key="idx"
                    :class="table.selected"
                    @click="$emit('select', idx)"
                >
                    <input
                        type="text"
                        :disabled="!table.editing"
                        :value="table.name"
                        @change="$emit('namechanged', idx, $event.target.value)"
                        @keydown.enter="$emit('locked', idx)"
                        @blur="$emit('locked', idx)">
                    <div class="table-name-control" v-if="!table.editing && !editable">
                        <i
                            class="fa fa-edit"
                            @click.stop="$emit('edit', idx)"
                        ></i><br>
                        <i
                            class="fa fa-trash-o"
                            @click.stop="$emit('delete', idx)"
                        ></i>
                    </div>
                    <div class="delete-modal" v-if="deleting === idx">
                        <div>真的要删除吗？</div>
                        <div class="delete-modal-but">
                            <i class="fa fa-check" @click.stop="$emit('deletecheck')"></i>
                            <i class="fa fa-remove" @click.stop="$emit('deletecancel')"></i>
                        </div>
                    </div>
                </div>
            </div>
            <div class="but-cont">
                <button
                    class="control-but"
                    @click="$emit('load')"
                    v-if="!editable"
                >
                    <i class="fa fa-file-excel-o"></i>
                </button>
                <button
                    class="control-but"
                    @click="$emit('add')"
                    v-if="!editable"
                >
                    <i class="fa fa-plus"></i>
                </button>
                <button class="control-but" @click="$emit('openwdata')"><i class="fa fa-warning"></i></button>
                <button class="control-but" @click="$emit('opensettings')"><i class="fa fa-gear"></i></button>
            </div>
            <img class="leftpad-logo" src="logo.png">
        </div>
    `,
};

const filterLeft = (val, str) => {
    let opr = '';
    if (str.slice(0, 2) === '>=' || str.slice(0, 2) === '<=' || str.slice(0, 2) === '==')
        opr = str.slice(0, 2);
    else if (str.slice(0, 1) === '>' || str.slice(0, 1) === '<')
        opr = str.slice(0, 1);
    const num = parseInt(str.slice(opr.length));
    if (opr === '') return true;
    if (isNaN(num)) return true;
    if (opr === '>=') return val >= num;
    if (opr === '<=') return val <= num;
    if (opr === '>') return val > num;
    if (opr === '<') return val < num;
    return val == num;
};

const chartPad = {
    props: [ 'data', 'cleanfilter', 'repairfilter', 'editable' ],
    emits: [ 'changed', 'delete', 'update:cleanfilter', 'update:repairfilter' ],
    computed: {
        fd() {
            const res = [];
            for (let rec of this.$props.data) {
                rec.filtered = true;
                if (rec.clean_left) {
                    if (!filterLeft(rec.clean_left, this.$props.cleanfilter))
                        rec.filtered = false;
                }
                if (rec.repair_left) {
                    if (!filterLeft(rec.repair_left, this.$props.repairfilter))
                        rec.filtered = false;
                }
                res.push(rec);
            }
            return res;
        }
    },
    methods: {},
    template: `
        <div class="chart-pad">
            <table>
                <thead>
                    <tr>
                        <th>序号</th>
                        <th>设备名称</th>
                        <th>清扫日期</th>
                        <th>清扫周期</th>
                        <th>结束日期</th>
                        <th><input
                            type="text"
                            placeholder="剩余天数"
                            class="left-filter"
                            :value="cleanfilter"
                            @input="$emit('update:cleanfilter', $event.target.value)">
                        </th>
                        <th>计表检修日期</th>
                        <th>计表检修周期</th>
                        <th>结束日期</th>
                        <th><input
                            type="text"
                            placeholder="剩余天数"
                            class="left-filter"
                            :value="repairfilter"
                            @input="$emit('update:repairfilter', $event.target.value)">
                        </th>
                        <th>照片</th>
                        <th>备注</th>
                        <th v-if="!editable"><i class="fa fa-trash-o"></i></th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="(record, idx) in fd" :key="idx">
                        <template v-if="record.filtered">
                        <td>{{ idx + 1 }}</td>
                        <td><input
                            :disabled="editable"
                            type="text"
                            @input="$emit('changed', idx, 'name', $event.target.value)"
                            :value="record.name">
                        </td>
                        <td>
                        
                        <input
                            :disabled="editable"
                            type="date"
                            :value="record.clean_start"
                            @input="$emit('changed', idx, 'clean_start', $event.target.value)"
                            style="width: 130px;"
                        ></td>
                        <td>
                            <select
                                :disabled="editable"
                                :value="record.clean_duration"
                                @change="$emit('changed', idx, 'clean_duration', $event.target.value)"
                                style="width: 50px;"
                            >
                                <option value="week">一周</option>
                                <option value="month">一月</option>
                                <option value="season">一季</option>
                                <option value="half">半年</option>
                                <option value="year">一年</option>
                                <option value="2year">两年</option>
                            </select>
                        </td>
                        <td>{{ record.clean_end }}</td>
                        <td :class="record.clean_left_class">{{ record.clean_left }}</td>
                        <td><input
                            :disabled="editable"
                            type="date"
                            :value="record.repair_start"
                            @input="$emit('changed', idx, 'repair_start', $event.target.value)"
                            style="width: 130px;"
                        ></td>
                        <td>
                            <select
                                :disabled="editable"
                                :value="record.repair_duration"
                                @change="$emit('changed', idx, 'repair_duration', $event.target.value)"
                                style="width: 50px;"
                            >
                                <option value="week">一周</option>
                                <option value="month">一月</option>
                                <option value="season">一季</option>
                                <option value="half">半年</option>
                                <option value="year">一年</option>
                                <option value="2year">两年</option>
                            </select>
                        </td>
                        <td>{{ record.repair_end }}</td>
                        <td :class="record.repair_left_class">{{ record.repair_left }}</td>
                        <td><input
                            :disabled="editable"
                            type="text"
                            @input="$emit('changed', idx, 'photo', $event.target.value)"
                            :value="record.photo">
                        </td>
                        <td><input
                            :disabled="editable"
                            type="text"
                            @input="$emit('changed', idx, 'note', $event.target.value)"
                            :value="record.note"
                            style="width: 70px;">
                        </td>
                        <td v-if="!editable"><i class="fa fa-trash-o" @click="$emit('delete', idx)"></i></td>
                        </template>
                    </tr>
                </tbody>
            </table>
        </div>
    `,
};

const warningDataPad = {
    props: [ 'data', 'sound', 'editable' ],
    emits: [ 'close', 'usercheck', 'usercheckall' ],
    data() {
        return {
            time_str: moment(new Date()).format("HH:mm")
        }
    },
    mounted() {
        setInterval(() => {
            this.time_str = moment(new Date()).format("HH:mm");
        }, 10000);
    },
    methods: {
        durChi2Eng(str) {
            if (str === "week") return "一周";
            else if (str === "month") return "一月";
            else if (str === "season") return "一季";
            else if (str === "half") return "半年";
            else if (str === "year") return "一年";
            else return "两年";
        }
    },
    template: `
        <div class="warning-data-pad">
            <i class="fa fa-remove wdata-close" @click="$emit('close')"></i>
            <div class="current-time">{{ time_str }}</div>
            <audio autoplay v-if="sound !== ''">
                <source :src="sound" type="audio/wav">
            </audio>
            <div>
                <div class="wdata-cont">
                    <table>
                        <thead>
                            <tr>
                                <th>表</th>
                                <th>设备名称</th>
                                <th>清扫周期</th>
                                <th>清扫结束日期</th>
                                <th>操作</th>
                                <th>计表检修周期</th>
                                <th>计表检修结束日期</th>
                                <th>操作</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr
                                v-for="(record, idx) in data"
                                :key="idx"
                            >
                                <td>{{ record.table }}</td>
                                <td>{{ record.name }}</td>
                                <td :class="record.clean_col">{{ record.clean_end }}</td>
                                <td>{{ durChi2Eng(record.clean_duration) }}</td>
                                <td>
                                    <i
                                        v-if="(record.clean_col === 'left-user-checked' || record.clean_col === 'left-admin-checked') && editable"
                                        class="fa fa-volume-up record-sbut"
                                        style="background-color: gray;"
                                        @click="$emit('usercheck', record.ti, record.ri, 0, 'clean')"
                                    ></i>
                                    <i
                                        v-if="record.clean_col === 'left-warning' || record.clean_col === 'left-error'"
                                        class="fa fa-volume-up record-sbut"
                                        style="background-color: red;"
                                        @click="$emit('usercheck', record.ti, record.ri, 0, 'user')"
                                    ></i>
                                    <i
                                        v-if="record.clean_col !== '' && record.clean_col !== 'left-admin-checked' && editable"
                                        class="fa fa-volume-up record-sbut"
                                        style="background-color: blue;"
                                        @click="$emit('usercheck', record.ti, record.ri, 0, 'admin')"
                                    ></i>
                                </td>
                                <td :class="record.repair_col">{{ record.repair_end }}</td>
                                <td>{{ durChi2Eng(record.repair_duration) }}</td>
                                <td>
                                    <i
                                        v-if="(record.repair_col === 'left-user-checked' || record.repair_col === 'left-admin-checked') && editable"
                                        class="fa fa-volume-up record-sbut"
                                        style="background-color: gray;"
                                        @click="$emit('usercheck', record.ti, record.ri, 1, 'clean')"
                                    ></i>
                                    <i
                                        v-if="record.repair_col === 'left-warning' || record.repair_col === 'left-error'"
                                        class="fa fa-volume-up record-sbut"
                                        style="background-color: red;"
                                        @click="$emit('usercheck', record.ti, record.ri, 1, 'user')"
                                    ></i>
                                    <i
                                        v-if="record.repair_col !== '' && record.repair_col !== 'left-admin-checked' && editable"
                                        class="fa fa-volume-up record-sbut"
                                        style="background-color: blue;"
                                        @click="$emit('usercheck', record.ti, record.ri, 1, 'admin')"
                                    ></i>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <div class="sound-check-cont">
                    <button
                        class="sound-check-but"
                        @click="$emit('usercheckall', 'user')"
                    >
                        <i class="fa fa-volume-up"></i> 普通用户
                    </button>
                    <button
                        v-if="editable"
                        class="sound-check-but"
                        @click="$emit('usercheckall', 'admin')"
                    >
                        <i class="fa fa-volume-up"></i> 管理者
                    </button>
                </div>
            </div>
        </div>
    `
}

const settingsPad = {
    props: [ ],
    data() {
        return {
            oldpass: this.$props.oldpass,
            newpass: this.$props.newpass,
            newpasscheck: this.$props.newpass,
            unsame: false,
            wrongpass: false,
        }
    },
    methods: {
        ok() {
            if (this.newpass === this.newpasscheck && this.newpass) {
                this.wrongpass = true;
                this.unsame = false;
                this.$emit('changepass', this.oldpass, this.newpass);
            }
            else {
                this.unsame = true;
            }
        },
    },
    emits: [ 'close', 'changepass' ],
    template: `
        <div class="settings-pad">
            <i class="fa fa-remove settings-close" @click="$emit('close')"></i>
            <div class="change-pass-cont">
                <input
                    :style="{ 'border-bottom': wrongpass ? '2px solid red' : '2px solid white' }"
                    v-model="oldpass"
                    class="password-input"
                    type="password"
                    placeholder="现在密码">
                <input
                    :style="{ 'border-bottom': unsame ? '2px solid red' : '2px solid white' }"
                    v-model="newpass"
                    class="password-input"
                    type="password"
                    placeholder="新密码">
                <input
                    :style="{ 'border-bottom': unsame ? '2px solid red' : '2px solid white' }"
                    v-model="newpasscheck"
                    class="password-input"
                    type="password"
                    placeholder="新密码确认">
                <button
                    class="password-button"
                    @click="ok"
                >
                    <i class="fa fa-check"></i> 确 认
                </button>
            </div>
        </div>
    `
};

const app = Vue.createApp(mainCont);
app.component('left-pad', leftPad);
app.component('chart-pad', chartPad);
app.component('login-pad', loginPad);
app.component('warning-data-pad', warningDataPad);
app.component('settings-pad', settingsPad);
app.directive('focus', {
    mounted(el) {
        el.focus()
    }
});
const vm = app.mount('#vue-container');

ipcRenderer.on('close', () => {
    fs.writeFileSync("dat", JSON.stringify({
        admin_password: vm.admin_password,
        normal_password: vm.normal_password,
        saved_admin_password: vm.saved_admin_password,
        saved_normal_password: vm.saved_normal_password,
        data: vm.data
    }));
});

ipcRenderer.on('read-path', (events, fpath) => {
    const dateStamp2Str = (str) => {
        return moment(parseInt(str - 25569) * 24 * 3600 * 1000 + 10).format("YYYY-MM-DD");
    }
    const getDur = (str) => {
        if (str === "一月") return "month";
        if (str === "半年") return "half";
        if (str === "两年") return "2year";
        if (str === "一周") return "week";
        return "year";
    }
    const sheets = xlsx.parse(fpath);
    for (let sheet of sheets) {
        const table_name = sheet.name;
        const sd = sheet.data;
        const td = [];
        for (let i = 2; i < sd.length; i++) {
            const r = sd[i];
            if (r.length == 0) break;
            td.push({
                name: r[2],
                clean_start: r[3] ? dateStamp2Str(r[3]) : "",
                clean_duration: r[4] ? getDur(r[4]) : "",
                repair_start: r[7] ? dateStamp2Str(r[7]) : "",
                repair_duration: r[4] ? getDur(r[4]) : "",
                photo: r[11],
                note: r[12]
            })
        }
        vm.data.push({
            name: table_name,
            data: td
        })
    }
});
