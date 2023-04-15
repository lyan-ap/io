<template>
  <div class="container">
    <div class="header">
      <select v-model="selected" placeholder="请选择" @change="selectExcel">
        <option
          v-for="option in options"
          :key="option.text"
          :value="option.value"
        >
          {{ option.text }}
        </option>
      </select>
      <div style="text-align: center" @click="update">
        <button class="update-button">财务报表计算</button>
      </div>
    </div>
    <div id="luckysheet"></div>
  </div>

  <Teleport to="body">
    <!-- use the modal component, pass in the prop -->
    <modal :show="showModal" @close="showModal = false">
      <template #header>
        <h2>财务报表计算结果</h2>
      </template>
      <template #body>
        <div>
          <div v-show="result.showSuccess" class="success">
            您的{{ result.success }};
          </div>
          <div v-show="result.showError" class="error">
            您的{{ result.error }}
          </div>
        </div>
      </template>
    </modal>
  </Teleport>
</template>

<script setup>
import { ref, onMounted } from 'vue';
import { exportExcel } from './export';
import Modal from './Modal.vue';
import LuckyExcel from 'luckyexcel';
const types = ['./assets/family-io.xlsx', './assets/family-o.xlsx'];
const showModal = ref(false);
const selected = ref(types[0]);
const result = ref({
  success: '',
  error: '',
});
const jsonData = ref({});
const options = ref([
  {
    text: '家庭收入支出表-年度.xlsx',
    value: types[0],
  },
  {
    text: '家庭资产负债表.xlsx',
    value: types[1],
  },
]);

function update() {
  const value = selected.value;
  const errors = [];
  const success = [];

  if (value === types[0]) {
    const [v1, v2, v3] = [3, 4, 5].map((r) => luckysheet.getCellValue(r, 7));
    const [i1, i2, i3] = ['结余比率', '财务负担比', '财务自由度'].map(
      (x) => `${x}指标`
    );
    if (v1 <= 0.2) {
      errors.push(`${i1}高于参考值`);
    } else {
      success.push(`${i1}合适`);
    }
    if (v2 >= 0.4) {
      errors.push(`${i2}低于参考值`);
    } else {
      success.push(`${i2}合适`);
    }
    if (v3 <= 1.0) {
      errors.push(`${i3}低于参考值`);
    } else {
      success.push(`${i3}美好`);
    }
  } else {
    const [v1, v2, v3] = [3, 4, 5].map((r) => luckysheet.getCellValue(r, 7));
    const [i1, i2, i3] = ['负债率', '投资比率', '流动性比率'].map(
      (x) => `${x}指标`
    );

    if (v1 >= 0.7) {
      errors.push(`${i1}高于参考值`);
    } else {
      success.push(`${i1}较安全`);
    }
    if (v2 <= 0.5) {
      errors.push(`${i1}低于参考值`);
    } else {
      success.push(`${i2}合适`);
    }
    if (v3 < 3.0 || v3 > 6) {
      errors.push(`${i1}不符合参考值`);
    } else {
      success.push(`${i3}适当`);
    }
  }
  const [successMsg, errorMsg] = [success, errors].map((x) => x.join(','));
  result.value = {
    success: successMsg,
    showSuccess: !!successMsg,
    showError: !!errorMsg,
    error: `${errorMsg},具体提升建议，请预约咨询！`,
  };
  showModal.value = true;
}

onMounted(() => {
  luckysheet.create({
    container: 'luckysheet',
    showinfobar: false,

    // showtoolbarConfig: {
    // showtoolbar: false,
    // showtoolbarConfig: {
    //   undoRedo: true,
    //   font: true,
    // },
    // },
    cellUpdated: (e) => {
      console.log(11, e);
    },
  });
  const { text, value } = options.value[0];
  loadTable(value, text);
  showModal.value = !true;
});

function selectExcel(evt) {
  const value = selected.value;
  const name = evt.target.options[evt.target.selectedIndex].innerText;
  loadTable(value, name);
}

function loadTable(value, name) {
  if (value == '') {
    return;
  }

  LuckyExcel.transformExcelToLuckyByUrl(
    value,
    name,
    (exportJson, luckysheetfile) => {
      if (exportJson.sheets == null || exportJson.sheets.length == 0) {
        alert(
          'Failed to read the content of the excel file, currently does not support xls files!'
        );
        return;
      }
      console.log('exportJson', exportJson);
      jsonData.value = exportJson;

      window.luckysheet.destroy();

      window.luckysheet.create({
        container: 'luckysheet',
        showinfobar: false,
        data: exportJson.sheets,
        title: exportJson.info.name,
        userInfo: exportJson.info.name.creator,
      });
    }
  );
}
</script>

<style scoped>
.container {
  padding: 0 8px;
}
.header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}
#luckysheet {
  margin: 0px;
  padding: 0px;
  position: absolute;
  width: 100%;
  left: 0px;
  top: 50px;
  bottom: 0px;
}

.update-button {
  width: 200px;
  height: 40px;
  border-width: 0px;
  border-radius: 3px;
  background: #1e90ff;
  cursor: pointer;
  outline: none;
  font-family: Microsoft YaHei;
  color: white;
  font-size: 17px;
}
.update-button:hover {
  background: #5599ff;
}

#tip {
  position: absolute;
  z-index: 1000000;
  left: 0px;
  top: 0px;
  bottom: 0px;
  right: 0px;
  background: rgba(255, 255, 255, 0.8);
  text-align: center;
  font-size: 40px;
  align-items: center;
  justify-content: center;
  display: flex;
}
.success {
  color: #1e90ff;
  margin-bottom: 16px;
}
.error {
  color: #f00;
  font-weight: 700;
}
</style>
