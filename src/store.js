import { createStore } from 'vuex';

export default createStore({
  state() {
    return {
      count: 0,
      excelA: [],
    };
  },
  mutations: {
    increment(state, data) {
      state.count = state.count + data;
    }
  },
  actions: {
    incrementAction({ commit }) {
      commit('increment');
    },
    handleImport({ commit }, data) {
      commit('increment', data);
    }
  },
  getters: {
    getCount: state => state.count
  }
});